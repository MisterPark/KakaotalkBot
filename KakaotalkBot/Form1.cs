using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
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
        private DateTime lastBotResetTime;

        public Form1(Bot bot)
        {
            lastBotResetTime = DateTime.Now;
            this.bot = bot;
            instance = this;
            InitializeComponent();

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 20;
            timer.Tick += Timer_Tick;
            timer.Start();

            WindowsMacro.Instance.Form = this;

            Settings.Save(Settings.Instance);
            textBox2.Text = Settings.Instance.ApplicationName;
            textBox3.Text = Settings.Instance.SpreadsheetId;

            Database.Instance.Initialize(Settings.Instance.ApplicationName, Settings.Instance.SpreadsheetId);

            Application.ApplicationExit += new EventHandler(OnApplicationExit);

            UpdateWindowList();

        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (bot.IsBotRunning)
            {
                button2.BackColor = Color.Green;
            }
            else
            {
                button2.BackColor = Color.Red;
            }

            label3.Text = lastBotResetTime.ToString("HH:mm:ss");
            label4.Text = DateTime.Now.ToString("HH:mm:ss");

            Point p = WindowsMacro.Instance.GetCursorPos();
            label5.Text = $"[{p.X}, {p.Y}]";

        }

        private void OnApplicationExit(object sender, EventArgs e)
        {

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

            UnregisterHotKey(this.Handle, 2);
            UnregisterHotKey(this.Handle, 1);
            Settings.Save(Settings.Instance);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bot.UpdateWindowList();
            bot.IsBotRunning = !bot.IsBotRunning;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            string roomName = listView1.SelectedItems[0].SubItems[0].Text;
            textBox1.Text = roomName;
            bot.TargetWindow = roomName;
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
            if (bot.IsBotRunning && string.IsNullOrEmpty(bot.TargetWindow) == false)
            {
                WindowsMacro.Instance.SendTextToChatroom(bot.TargetWindow, $"o");
                bot.IsBotRunning = false;
                //Thread.Sleep(5000);
                WindowsMacro.Instance.CloseChatRoom(bot.TargetWindow);
                Thread.Sleep(3000);


                WindowsMacro.Instance.OpenChatRoom(bot.TargetWindow);
                //Thread.Sleep(5000);
                //bot.UpdateWindowList();
                //bot.IsBotRunning = true;
                //WindowsMacro.Instance.SendTextToChatroom(bot.TargetWindow, $"[시스템] 코몽봇 껐켰 테스트 종료");
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
                    WindowsMacro.Instance.SetCursor(x, y);
                    WindowsMacro.Instance.ClickLeft();
                    Thread.Sleep(50);
                    WindowsMacro.Instance.ClickLeft();
                    Thread.Sleep(50);

                    
                });

                ScreenPixelDetector.Instance.Start(x, y);
            }


        }
    }
}
