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
            bot.TargetWindow = "";
            bot.IsBotRunning = false;
            bot.VoiceRoomBot.IsBotRunning = false;

            //VoiceRoomBot voiceRoomBot = new VoiceRoomBot();
            //voiceRoomBot.TargetWindow = "흑우방";


            DateTime utc = GetUtc();


            var kstZone = TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time");
            DateTime kst = TimeZoneInfo.ConvertTimeFromUtc(utc, kstZone);

            //MessageBox.Show("KST : " + kst.ToString("yyyy-MM-dd HH:mm:ss"));

            DateTime limit = new DateTime(2026, 2, 28, 0, 0, 0, 0, DateTimeKind.Local);
            var t = limit - kst;

            //if(t.TotalDays < 0)
            //{
            //    MessageBox.Show("사용 가능 기간 초과");
            //    return;
            //}


            Form1 form = new Form1(bot, bot.VoiceRoomBot);
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

                    bot.Update();
                    //voiceRoomBot.Update();
                    // 입력·수신은 별도 작업자에서 실행하므로 유휴 루프의 과도한 할당을 제한합니다.
                    System.Threading.Thread.Sleep(10);
                }
            }
            if (ScreenPixelDetector.Instance.IsRunning) ScreenPixelDetector.Instance.Stop();
            bot.VoiceRoomBot.Dispose();
            StaInputWorker.Instance.Complete();
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
