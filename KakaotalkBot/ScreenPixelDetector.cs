using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;

namespace KakaotalkBot
{
    public class ScreenPixelDetector
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("gdi32.dll")]
        static extern uint GetPixel(IntPtr hdc, int x, int y);

        struct POINT
        {
            public int X;
            public int Y;
        }

        private static ScreenPixelDetector instance;
        public static ScreenPixelDetector Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ScreenPixelDetector();
                }
                return instance;
            }
        }

        private Thread t;
        private List<Action> actions = new List<Action>();

        private Color oldColor = new Color();

        public bool IsRunning { get; set; } = false;

        public void Start(int x, int y)
        {
            if (t == null)
            {
                t = new Thread(() =>
                {
                    oldColor = GetScreenPixelColor(x, y);

                    while (true)
                    {
                        Color color = GetScreenPixelColor(x, y);
                        if(oldColor != color)
                        {
                            oldColor = color;
                            Invoke();
                        }
                    }
                });

                t.IsBackground = true; // 프로그램 종료 시 자동 종료

            }
            else
            {
                if (t.ThreadState == ThreadState.Running)
                {
                    return;
                }
            }

            t.Start();

            IsRunning = true;
        }

        public void Stop()
        {
            if (t != null)
            {
                t.Abort();
            }

            IsRunning = false;
        }

        public void AddListener(Action action)
        {
            actions.Add(action);
        }

        public void RemoveListener(Action action)
        {
            actions.Remove(action);
        }

        public void Invoke()
        {
            int count = actions.Count;
            for (int i = 0; i < count; i++)
            {
                actions[i].Invoke();
            }
        }

        public static Color GetScreenPixelColor(int x, int y)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            uint pixel = GetPixel(hdc, x, y);
            ReleaseDC(IntPtr.Zero, hdc);

            return Color.FromArgb(
                (int)(pixel & 0x000000FF),
                (int)(pixel & 0x0000FF00) >> 8,
                (int)(pixel & 0x00FF0000) >> 16
            );
        }
    }
}
