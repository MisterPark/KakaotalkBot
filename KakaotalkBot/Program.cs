using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KakaotalkBot
{
    internal static class Program
    {
        public static bool isThreadRunning = false;
        public static bool isBotRunning = false;
        public static bool ShutdownFlag = false;

        [STAThread]
        static void Main()
        {
            Time.Initialize();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Bot bot = new Bot();
            bot.TargetWindow = "흑우방";
            CustomTimer rebootTimer = new CustomTimer(14400000);

            VoiceRoomBot voiceRoomBot = new VoiceRoomBot();


            DateTime utc = GetUtc();


            var kstZone = TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time");
            DateTime kst = TimeZoneInfo.ConvertTimeFromUtc(utc, kstZone);

            //MessageBox.Show("KST : " + kst.ToString("yyyy-MM-dd HH:mm:ss"));

            DateTime limit = new DateTime(2026, 2, 28, 0, 0, 0, 0, DateTimeKind.Local);
            var t = limit - kst;

            if(t.TotalDays < 0)
            {
                MessageBox.Show("사용 가능 기간 초과");
                return;
            }


            Form1 form = new Form1(bot, voiceRoomBot);
            form.Show();

            NativeMessage msg;
            while (!ShutdownFlag)
            {
                if (PeekMessage(out msg, IntPtr.Zero, 0, 0, 1))
                {
                    if (msg.msg == 0x0012) // WM_QUIT
                        break;

                    TranslateMessage(ref msg);
                    DispatchMessage(ref msg);
                }
                else
                {
                    Time.Update();

                    if(StaticVariable.AutoReboot)
                    {
                        if (string.IsNullOrEmpty(bot.TargetWindow) == false)
                        {
                            if (WindowsMacro.Instance.IsKakaoTalkOpen() == false)
                            {
                                WindowsMacro.Instance.OpenChatRoom(bot.TargetWindow);
                                Thread.Sleep(1000);
                                WindowsMacro.Instance.SendTextToChatroom(bot.TargetWindow, "[시스템] 코몽봇이 다시 시작되었습니다.");
                                bot.Start();
                            }

                            if (WindowsMacro.Instance.IsChatRoomOpen(bot.TargetWindow) == false)
                            {
                                WindowsMacro.Instance.OpenChatRoom(bot.TargetWindow);
                                Thread.Sleep(1000);
                                WindowsMacro.Instance.SendTextToChatroom(bot.TargetWindow, "[시스템] 코몽봇이 다시 시작되었습니다.");
                                bot.Start();
                            }
                        }

                        if (rebootTimer.Check(Time.DeltaTime) && bot.IsBotRunning)
                        {
                            WindowsMacro.Instance.SendTextToChatroom(bot.TargetWindow, "[시스템] 원활한 사용을 위해 봇이 재기동됩니다.\n(1분 정도 소요됨.)");
                            bot.Stop();
                            WindowsMacro.Instance.CloseChatRoom(bot.TargetWindow);
                        }
                    }
                    
                    bot.Update();
                    voiceRoomBot.Update();
                    System.Threading.Thread.Sleep(0);
                }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct NativeMessage
        {
            public IntPtr handle;
            public uint msg;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public System.Drawing.Point p;
        }

        [DllImport("user32.dll")]
        public static extern bool PeekMessage(out NativeMessage message, IntPtr handle,
            uint filterMin, uint filterMax, uint remove);

        [DllImport("user32.dll")]
        public static extern bool TranslateMessage([In] ref NativeMessage message);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage([In] ref NativeMessage message);

        private static readonly HttpClient _http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(3)
        };
        public static DateTime GetUtc()
        {
            using (var req = new HttpRequestMessage(HttpMethod.Head, "https://www.naver.com"))
            using (var resp = _http
                .SendAsync(req, HttpCompletionOption.ResponseHeadersRead)
                .GetAwaiter()
                .GetResult())
            {
                if (!resp.Headers.Date.HasValue)
                    throw new Exception("Naver Date header not found.");

                // Date 헤더는 UTC
                return resp.Headers.Date.Value.UtcDateTime;
            }
        }
    }
}
