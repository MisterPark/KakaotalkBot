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

        private static readonly ScreenPixelDetector instance = new ScreenPixelDetector();
        public static ScreenPixelDetector Instance { get { return instance; } }
        private readonly object gate = new object();
        private readonly List<Action> actions = new List<Action>();
        private readonly Func<int, int, Color> readPixel;
        private CancellationTokenSource cancellation;
        private long generation;
        private volatile bool running;
        private volatile string lastError;
        public bool IsRunning { get { return running; } }
        public string LastError { get { return lastError; } }
        internal long Generation { get { return Interlocked.Read(ref generation); } }
        internal bool IsCurrentSession(long value) { return running && Generation == value; }
        public ScreenPixelDetector() : this(GetScreenPixelColor) { }
        internal ScreenPixelDetector(Func<int, int, Color> readPixel) { this.readPixel = readPixel; }

        public void Start(int x, int y)
        {
            lock (gate)
            {
                if (running) return;
                var source = new CancellationTokenSource();
                cancellation = source;
                long session = Interlocked.Increment(ref generation);
                running = true; lastError = null;
                var thread = new Thread(() => Run(x, y, source, session)) { IsBackground = true, Name = "화면 픽셀 감지" };
                thread.Start();
            }
        }

        private void Run(int x, int y, CancellationTokenSource source, long session)
        {
            try
            {
                Color oldColor = readPixel(x, y);
                // 초당 최대 20회 검사합니다. 변화가 없을 때 CPU를 계속 점유하지 않습니다.
                while (!source.Token.WaitHandle.WaitOne(50))
                {
                    Color color = readPixel(x, y);
                    if (source.IsCancellationRequested) break;
                    if (oldColor != color)
                    {
                        oldColor = color;
                        Action[] snapshot;
                        lock (gate) snapshot = actions.ToArray();
                        foreach (var action in snapshot)
                        {
                            if (!IsCurrentSession(session)) break;
                            action();
                        }
                    }
                }
            }
            catch (Exception error)
            {
                lock (gate) if (Generation == session) lastError = "픽셀 감지 중지: " + error.Message;
            }
            finally
            {
                lock (gate)
                {
                    if (Generation == session) { running = false; cancellation = null; }
                    source.Dispose();
                }
            }
        }

        public void Stop()
        {
            lock (gate)
            {
                running = false;
                Interlocked.Increment(ref generation);
                if (cancellation != null) cancellation.Cancel();
                cancellation = null;
            }
        }

        public void AddListener(Action action)
        {
            if (action == null) throw new ArgumentNullException("action");
            lock (gate) if (!actions.Contains(action)) actions.Add(action);
        }
        public void RemoveListener(Action action) { lock (gate) actions.Remove(action); }
        public void Invoke()
        {
            Action[] snapshot;
            lock (gate) snapshot = actions.ToArray();
            foreach (var action in snapshot) action();
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
