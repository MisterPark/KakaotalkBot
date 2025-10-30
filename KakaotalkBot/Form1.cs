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
        private static Form1 instance;
        public static Form1 Instance
        {
            get { return instance; }
        }
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


        private Queue<Command> commands = new Queue<Command>();

        private List<string> chatLog = new List<string>();
        string lastChat = "xx";
       

        private DateTime lastUpdate = DateTime.MinValue;

        CustomTimer soliloquyTimer = new CustomTimer(300000);
        CustomTimer newsTimer = new CustomTimer(3600000);
        CustomTimer dbTimer = new CustomTimer(60000);

        private Queue<QuizAnswer> quizAnswers = new Queue<QuizAnswer>();
        private bool isCorrect = false;

        private System.Windows.Forms.Timer timer;

        public Form1()
        {
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

            isThreadRunning = false;
            thread.Join();

            UnregisterHotKey(this.Handle, 2);
            UnregisterHotKey(this.Handle, 1);
            Settings.Save(Settings.Instance);
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
