using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KakaotalkBot
{
    internal static class Program
    {
        public static bool isThreadRunning = false;
        public static bool isBotRunning = false;

        [STAThread]
        static void Main()
        {
            Time.Initialize();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Form1 form = new Form1();
            form.Show();

            Bot bot = new Bot();

            NativeMessage msg;
            while (true)
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



    }
}
