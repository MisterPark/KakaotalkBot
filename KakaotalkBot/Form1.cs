using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KakaotalkBot
{
    public partial class Form1 : Form
    {
        private static Form1 instance;
        public static Form1 Instance
        {
            get { return instance; }
        }
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

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


        private System.Windows.Forms.Timer timer;
        private Bot bot;
        private VoiceRoomBot voiceRoomBot;
        private DateTime lastBotResetTime;
        private readonly CancellationTokenSource catalogCancellation = new CancellationTokenSource();
        private Task<ChatCatalog> catalogTask;
        private ChatCatalog catalog;
        private ComboBox accountSelector;
        private TextBox roomSearch;
        private TextBox ownIdInput;
        private Label databaseStatus;
        private bool closing, closeReady;
        private string databaseError;

        public Form1(Bot bot, VoiceRoomBot voiceRoomBot)
        {
            lastBotResetTime = DateTime.Now;
            this.bot = bot;

            instance = this;
            InitializeComponent();
            InitializeDatabaseControls();
            InitializeOperationsControls();

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 100;
            timer.Tick += Timer_Tick;
            timer.Start();

            WindowsMacro.Instance.Form = this;

            Settings.Save(Settings.Instance);
            textBox2.Text = Settings.Instance.ApplicationName;
            textBox3.Text = Settings.Instance.SpreadsheetId;

            try { Database.Instance.Initialize(Settings.Instance.ApplicationName, Settings.Instance.SpreadsheetId); }
            catch (Exception error) { databaseError = "DB 초기화 실패: " + error.GetBaseException().Message; }

            Application.ApplicationExit += new EventHandler(OnApplicationExit);

            this.voiceRoomBot = voiceRoomBot;
            Shown += (sender, args) => RefreshRoomCatalog();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            TickOperations();
            if (catalogTask != null && catalogTask.IsCompleted)
            {
                var completed = catalogTask;
                catalogTask = null;
                if (completed.IsFaulted) databaseError = "방 목록 읽기 실패: " + completed.Exception.GetBaseException().Message;
                else if (!completed.IsCanceled)
                {
                    catalog = completed.Result;
                    accountSelector.Items.Clear();
                    accountSelector.Items.Add("모든 계정");
                    foreach (string account in catalog.Rooms.Select(r => r.AccountPath).Distinct()) accountSelector.Items.Add(Path.GetFileName(account));
                    accountSelector.SelectedIndex = 0;
                    databaseStatus.Text = catalog.Rooms.Count + "개 방 · " + string.Join(" / ", catalog.Errors);
                    foreach (ListViewItem item in listView1.Items)
                    {
                        var room = item.Tag as ChatRoomInfo;
                        if (room != null && room.AccountPath == Settings.Instance.ChatAccountPath && room.ChatId == Settings.Instance.ChatRoomId)
                        { item.Selected = true; item.EnsureVisible(); break; }
                    }
                }
            }
            bool available = !closing && !bot.IsBotRunning && !bot.IsReceiverStopping && catalogTask == null;
            listView1.Enabled = accountSelector.Enabled = roomSearch.Enabled = ownIdInput.Enabled = textBox1.Enabled = available;
            button3.Enabled = available;
            button2.Enabled = !bot.IsReceiverStopping && catalogTask == null && bot.SelectedRoom != null;
            button2.Text = bot.IsBotRunning ? "DB 수신 중지" : bot.IsReceiverStopping ? "수신 종료 중…" : "DB 수신 시작";
            if (bot.HasReceiver && catalogTask == null) databaseStatus.Text = bot.ReceiveStatus;
            if (News.LastRefreshError != null) databaseStatus.Text = News.LastRefreshError;
            if (Database.Instance.ContentRefreshError != null) databaseStatus.Text = Database.Instance.ContentRefreshError;
            if (databaseError != null) databaseStatus.Text = databaseError;
            if (bot.RoomRecycleStatus != null) databaseStatus.Text = bot.RoomRecycleStatus;
            if (bot.LastSendError != null) databaseStatus.Text = bot.LastSendError;
            if (bot.LastProcessingError != null) databaseStatus.Text = bot.LastProcessingError;
            if (bot.IsBotRunning)
            {
                button2.BackColor = Color.Green;
            }
            else
            {
                button2.BackColor = Color.Red;
            }

            if (voiceRoomBot.IsBotRunning)
            {
                button1.BackColor = Color.Green;
            }
            else
            {
                button1.BackColor = Color.Red;
            }

            label3.Text = (bot.LastRoomRecycle == DateTime.MinValue ? lastBotResetTime : bot.LastRoomRecycle).ToString("HH:mm:ss");
            label4.Text = DateTime.Now.ToString("HH:mm:ss");

            Point p = WindowsMacro.Instance.GetCursorPos();
            label5.Text = $"[{p.X}, {p.Y}]";
            label7.Text = $"[{p.X}, {p.Y}]";

            if(voiceRoomBot.CurrentScreen != null)
            {
                pictureBox1.Size = new Size(voiceRoomBot.CurrentScreen.Width, voiceRoomBot.CurrentScreen.Height);
                pictureBox1.Image = voiceRoomBot.CurrentScreen;
            }

            if (voiceRoomBot.CurrentScreen2 != null)
            {
                pictureBox2.Size = new Size(voiceRoomBot.CurrentScreen2.Width, voiceRoomBot.CurrentScreen2.Height);
                pictureBox2.Image = voiceRoomBot.CurrentScreen2;
            }

            if (voiceRoomBot.CurrentScreen3 != null)
            {
                pictureBox3.Size = new Size(voiceRoomBot.CurrentScreen3.Width, voiceRoomBot.CurrentScreen3.Height);
                pictureBox3.Image = voiceRoomBot.CurrentScreen3;
            }

        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            catalogCancellation.Cancel();
            bot.Stop();
            Program.ShutdownFlag = true;
        }

        private void InitializeDatabaseControls()
        {
            // 디자이너가 컨테이너를 만들지 않은 경우 직접 생성하여 툴팁도 폼과 함께 해제합니다.
            if (components == null) components = new System.ComponentModel.Container();

            tabPage2.Text = "채팅 DB";
            button3.Text = "DB 목록 읽기";
            button3.SetBounds(8, 33, 104, 23);
            accountSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            accountSelector.SetBounds(118, 33, 304, 23);
            accountSelector.SelectedIndexChanged += (sender, args) => FilterRooms();
            roomSearch = new TextBox();
            roomSearch.SetBounds(8, 62, 414, 21);
            roomSearch.TextChanged += (sender, args) => FilterRooms();
            listView1.SetBounds(8, 90, 414, 154);
            listView1.FullRowSelect = true;
            listView1.MultiSelect = false;
            columnHeader1.Text = "채팅방";
            columnHeader1.Width = 235;
            columnHeader2.Text = "방 ID";
            columnHeader2.Width = 155;
            databaseStatus = new Label { AutoSize = false, Text = "DB 목록을 읽을 준비 중입니다." };
            databaseStatus.SetBounds(8, 611, 796, 38);
            var ownLabel = new Label { AutoSize = false, Text = "본인 작성자 ID (자동 확인 실패 시 입력)" };
            ownLabel.SetBounds(428, 550, 365, 20);
            ownIdInput = new TextBox { Text = Settings.Instance.ChatOwnAuthorId == 0 ? "" : Settings.Instance.ChatOwnAuthorId.ToString() };
            ownIdInput.SetBounds(428, 574, 365, 21);
            var outputLabel = new Label { AutoSize = false, Text = "저장 위치: " + ChatCatalog.OutputRoot };
            outputLabel.SetBounds(8, 654, 796, 36);
            tabPage2.Controls.AddRange(new Control[] { accountSelector, roomSearch, databaseStatus, ownLabel, ownIdInput, outputLabel });
            var tips = new ToolTip(components);
            tips.SetToolTip(textBox1, "답변을 보낼 카카오톡 채팅창 이름입니다. DB의 방 이름과 다르면 수정하세요.");
            tips.SetToolTip(roomSearch, "방 이름 또는 방 ID로 검색합니다.");
            tips.SetToolTip(accountSelector, "복호화 가능한 계정 목록입니다. 현재 로그인 계정이라고 단정하지 않습니다.");
            richTextBox1.ReadOnly = true;
        }


        private void RefreshRoomCatalog()
        {
            if (catalogTask != null || bot.IsBotRunning || bot.IsReceiverStopping) return;
            databaseError = null;
            bot.SelectRoom(null);
            textBox1.Clear();
            databaseStatus.Text = "계정과 채팅방 DB 목록을 읽는 중…";
            catalogTask = Task.Run(() => ChatCatalog.Load(catalogCancellation.Token), catalogCancellation.Token);
        }

        private void FilterRooms()
        {
            listView1.BeginUpdate();
            listView1.Items.Clear();
            if (catalog != null)
            foreach (var room in catalog.Rooms.OrderBy(r => r.Name, StringComparer.CurrentCultureIgnoreCase).ThenBy(r => r.ChatId))
            {
                if (string.Equals(room.Name?.Trim(), "(알 수 없음)", StringComparison.Ordinal)) continue;
                if (accountSelector.SelectedIndex > 0 && Path.GetFileName(room.AccountPath) != Convert.ToString(accountSelector.SelectedItem)) continue;
                if ((room.Name + " " + room.ChatId).IndexOf(roomSearch.Text.Trim(), StringComparison.OrdinalIgnoreCase) < 0) continue;
                listView1.Items.Add(new ListViewItem(new[] { room.Name, room.ChatId.ToString() }) { Tag = room });
            }
            listView1.EndUpdate();
        }

        public void UpdateChatLog(string text)
        {
            richTextBox1.Text = text;
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.ScrollToCaret();
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
                    if (voiceRoomBot.IsBotRunning)
                    {
                        voiceRoomBot.Stop();
                    }
                    else
                    {
                        int x = (int)numericUpDown3.Value;
                        int y = (int)numericUpDown4.Value;
                        int delay = (int)numericUpDown5.Value;
                        voiceRoomBot.Start(x, y, delay);
                    }
                }

                if (id == 2)
                {

                }
            }
            base.WndProc(ref m);
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            if(voiceRoomBot.IsBotRunning)
            {
                voiceRoomBot.Stop();
            }
            else
            {
                int x = (int)numericUpDown3.Value;
                int y = (int)numericUpDown4.Value;
                int delay = (int)numericUpDown5.Value;
                voiceRoomBot.Start(x, y, delay);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, 1, 0, (int)Keys.F5);
            RegisterHotKey(this.Handle, 2, 0, (int)Keys.F6);
        }

        private async void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!closeReady)
            {
                e.Cancel = true;
                if (closing) return;
                closing = true;
                timer.Stop();
                Enabled = false;
                bot.Stop();
                if (bot.HasPendingUserSave)
                {
                    // 저장되지 않은 경험치·방 상태·이력을 버리지 않도록 재시도 기회를 남긴다.
                    closing = false;
                    Enabled = true;
                    timer.Start();
                    databaseStatus.Text = bot.LastProcessingError;
                    return;
                }
                catalogCancellation.Cancel();
                databaseStatus.Text = "수신 작업을 종료하고 키를 해제하는 중…";
                try
                {
                    if (catalogTask != null) await catalogTask;
                    await bot.ReceiverCompletion;
                }
                catch (Exception) { }
                closeReady = true;
                Close();
                return;
            }
            timer.Stop();

            UnregisterHotKey(this.Handle, 2);
            UnregisterHotKey(this.Handle, 1);
            Settings.Save(Settings.Instance);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            databaseError = null;
            try
            {
                if (bot.IsBotRunning) { bot.Stop(); return; }
                long ownId = 0;
                if (!string.IsNullOrWhiteSpace(ownIdInput.Text) && (!long.TryParse(ownIdInput.Text, out ownId) || ownId <= 0))
                    throw new InvalidOperationException("본인 작성자 ID는 양의 정수로 입력하거나 비워 주세요.");
                bot.OwnAuthorId = ownId;
                bot.TargetWindow = textBox1.Text.Trim();
                Settings.Instance.ChatOwnAuthorId = ownId;
                Settings.Instance.ChatSendRoomName = bot.TargetWindow;
                Settings.Save(Settings.Instance);
                bot.Start();
            }
            catch (Exception ex) { databaseError = ex.Message; databaseStatus.Text = ex.Message; }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            var room = listView1.SelectedItems[0].Tag as ChatRoomInfo;
            if (room == null || bot.IsBotRunning || bot.IsReceiverStopping) return;
            databaseError = null;
            bool restored = Settings.Instance.ChatAccountPath == room.AccountPath && Settings.Instance.ChatRoomId == room.ChatId;
            if (Settings.Instance.ChatAccountPath != room.AccountPath)
            {
                ownIdInput.Text = "";
                Settings.Instance.ChatOwnAuthorId = 0;
            }
            bot.SelectRoom(room);
            textBox1.Text = restored && !string.IsNullOrWhiteSpace(Settings.Instance.ChatSendRoomName) ? Settings.Instance.ChatSendRoomName : room.Name;
            Settings.Instance.ChatAccountPath = room.AccountPath;
            Settings.Instance.ChatRoomId = room.ChatId;
            databaseStatus.Text = room.Name + " · 방 " + room.ChatId + " · " + Path.GetFileName(room.AccountPath);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //SendTextToChatroom(textBox1.Text, $"앙 기모띠");
            RefreshRoomCatalog();
        }


        private void button6_Click(object sender, EventArgs e)
        {
            if (bot.IsBotRunning && string.IsNullOrEmpty(bot.TargetWindow) == false)
            {
                string room = bot.TargetWindow;
                StaInputWorker.Instance.Invoke(() =>
                {
                    WindowsMacro.Instance.CloseChatRoom(room);
                    Thread.Sleep(3000);
                    WindowsMacro.Instance.OpenChatRoom(room);
                });
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            WindowsMacro.Instance.CloseChatRoom("흑우방");
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int x = (int)numericUpDown1.Value;
            int y = (int)numericUpDown2.Value;

            if (ScreenPixelDetector.Instance.IsRunning)
            {
                ScreenPixelDetector.Instance.Stop();
            }
            else
            {

                ScreenPixelDetector.Instance.AddListener(() =>
                {
                    StaInputWorker.Instance.Invoke(() =>
                    {
                        if (!ScreenPixelDetector.Instance.IsRunning) return;
                        WindowsMacro.Instance.SetCursor(x, y);
                        WindowsMacro.Instance.ClickLeft();
                        Thread.Sleep(50);
                        WindowsMacro.Instance.ClickLeft();
                        Thread.Sleep(50);
                    });
                });

                ScreenPixelDetector.Instance.Start(x, y);
            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            voiceRoomBot.IsClickMacroRunning =
            checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            StaticVariable.AutoReboot = checkBox2.Checked;
        }
    }
}
