using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;

namespace KakaotalkBot
{
    public class VoiceRoomBot : IDisposable
    {
        private volatile bool isBotRunning = false;
        private long lastInputTick;
        private readonly AutoResetEvent recognitionReady = new AutoResetEvent(false);
        private readonly object lifecycleGate = new object();
        private Thread recognitionThread;
        private readonly Action recognitionAction;
        private volatile bool disposed;
        private volatile string recognitionError;
        private volatile string inputStatus;
        public string InputStatus { get { return inputStatus; } }
        public string RecognitionError { get { return recognitionError; } }
        private int clickPending;
        private bool clickProposed;
        private long cycleSession;
        private long presenterOpened;
        private string targetWindow = string.Empty;
        private Frame frame1, frame2, frame3;
        private sealed class Frame
        {
            internal IntPtr Handle;
            internal Point Position, Size, Origin;
            internal long Tick, Session;
            internal string Target;
            internal bool Voice;
        }
        private static long Tick { get { return System.Diagnostics.Stopwatch.GetTimestamp(); } }
        private static long Age(long tick) { return (Tick - tick) * 1000 / System.Diagnostics.Stopwatch.Frequency; }
        private long inputSession;
        private readonly object screenGate = new object();
        private long screenVersion;
        private CustomTimer autoClickTimer;
        private CustomTimer screenCaptureTimer = new CustomTimer(100);
        private CustomTimer screenCaptureTimer2 = new CustomTimer(90);
        private CustomTimer screenCaptureTimer3 = new CustomTimer(1000);
        private CustomTimer autoPresenterTimer = new CustomTimer(10000);
        private CustomTimer lineDetectorTimer = new CustomTimer(10000);
        private CustomTimer acceptTimer = new CustomTimer(90);
        private CustomTimer joinTimer = new CustomTimer(3000);
        private CustomTimer voiceRoomExitTimer = new CustomTimer(3000);

        private int modifiedY = 560;
        private Bitmap line;
        private Bitmap manager;
        private Bitmap manager2;
        private Bitmap host;
        private Bitmap accept;
        private Bitmap accept2;
        private Bitmap join;
        private Bitmap voiceRoomExit;
        private Bitmap check;

        public int X { get; set; } = 0;
        public int Y { get; set; } = 0;
        public bool IsClickMacroRunning { get; set; } = false;
        public bool IsBotRunning
        {
            get { return isBotRunning; }
            set
            {
                lock (lifecycleGate)
                {
                    if (disposed) return;
                    if (isBotRunning != value) Interlocked.Increment(ref inputSession);
                    isBotRunning = value;
                    Interlocked.Exchange(ref lastInputTick, 0);
                    if (value && recognitionThread == null)
                    {
                        recognitionThread = new Thread(RecognitionLoop) { IsBackground = true, Name = "보이스룸 화면 인식" };
                        recognitionThread.Start();
                    }
                    recognitionReady.Set();
                }
            }
        }
        public Bitmap CurrentScreen { get; private set; }
        public Bitmap CurrentScreen2 { get; private set; }
        public Bitmap CurrentScreen3 { get; private set; }
        public string TargetWindow
        {
            get { return targetWindow; }
            set
            {
                if (targetWindow == value) return;
                targetWindow = value;
                Interlocked.Increment(ref inputSession);
            }
        }


        // 테스트에서는 실제 화면과 입력 장치 없이 작업자 수명과 분리를 검증합니다.
        internal VoiceRoomBot(Action recognitionAction) { this.recognitionAction = recognitionAction; }

        public VoiceRoomBot() : this(null)
        {
            line = new Bitmap("리스너경계선.bmp");
            manager = new Bitmap("방장.bmp");
            manager2 = new Bitmap("부방장.bmp");
            host = new Bitmap("진행자.bmp");
            accept = new Bitmap("수락.bmp");
            accept2 = new Bitmap("수락2.bmp");
            join = new Bitmap("참여.bmp");
            voiceRoomExit = new Bitmap("보룸닫기.bmp");
            check = new Bitmap("확인.bmp");
        }

        private void RecognitionLoop()
        {
            try
            {
                while (!disposed)
                {
                    recognitionReady.WaitOne(IsBotRunning ? 90 : Timeout.Infinite);
                    if (disposed) break;
                    if (!IsBotRunning || Volatile.Read(ref clickPending) != 0) continue;
                    try
                    {
                        lock (screenGate)
                        {
                            if (recognitionAction != null) recognitionAction(); else Recognize();
                        }
                        recognitionError = null;
                    }
                    catch (Exception ex) { recognitionError = ex.GetType().Name + ": " + ex.Message; }
                }
            }
            finally
            {
                lock (screenGate)
                {
                    foreach (var bitmap in new[] { CurrentScreen, CurrentScreen2, CurrentScreen3,
                        line, manager, manager2, host, accept, accept2, join, voiceRoomExit, check })
                        if (bitmap != null) bitmap.Dispose();
                    CurrentScreen = CurrentScreen2 = CurrentScreen3 = null;
                    screenVersion++;
                }
                lock (lifecycleGate) recognitionReady.Dispose();
            }
        }

        private void Recognize()
        {
            clickProposed = false;
            cycleSession = Interlocked.Read(ref inputSession);
            long tick = Tick;
            long previous = Interlocked.Exchange(ref lastInputTick, tick);
            long elapsed = previous == 0 ? 0 : (tick - previous) * 1000 / System.Diagnostics.Stopwatch.Frequency;
            if (previous == 0) Interlocked.Exchange(ref presenterOpened, 0);
            if (screenCaptureTimer.Check(elapsed)) ProcessCaptureScreen();
            if (screenCaptureTimer2.Check(elapsed)) ProcessCaptureScreen2();
            if (screenCaptureTimer3.Check(elapsed)) ProcessCaptureScreen3();
            long opened = Interlocked.Read(ref presenterOpened);
            if (opened != 0)
            {
                if (Age(opened) > 1500) Interlocked.Exchange(ref presenterOpened, 0);
                else if (frame1 != null && frame1.Tick > opened) ProcessAutoPresenter2();
                return;
            }
            if (acceptTimer.Check(elapsed)) ProcessAccept();
            if (joinTimer.Check(elapsed)) ProcessJoin();
            if (voiceRoomExitTimer.Check(elapsed)) ProcessVoiceRoomExit();
            if (lineDetectorTimer.Check(elapsed)) ProcessLineDetect();
            if (autoPresenterTimer.Check(elapsed)) ProcessAutoPresenter();
            if (autoClickTimer != null && autoClickTimer.Check(elapsed) && IsClickMacroRunning)
                QueueClick(null, new Point(X, Y), false, false);
        }

        public void Dispose()
        {
            lock (lifecycleGate)
            {
                if (disposed) return;
                isBotRunning = false;
                Interlocked.Increment(ref inputSession);
                disposed = true;
                recognitionReady.Set();
                if (recognitionThread == null)
                {
                    // 시작하지 않은 인스턴스도 동일한 정리 경로를 사용합니다.
                    recognitionThread = new Thread(RecognitionLoop) { IsBackground = true };
                    recognitionThread.Start();
                }
            }
        }

        // 화면 인식 결과는 한 건만 대기시킵니다. 커서 이동과 클릭은 하나의 STA 작업입니다.
        private bool QueueClick(Frame frame, Point point, bool right, bool presenter)
        {
            long session = cycleSession;
            long queued = Tick;
            if (clickProposed || Interlocked.CompareExchange(ref clickPending, 1, 0) != 0) return false;
            clickProposed = true;
            bool posted = StaInputWorker.Instance.TryPostBackground(this, () =>
            {
                try
                {
                    if (disposed || !IsBotRunning || session != Interlocked.Read(ref inputSession)) return;
                    if (frame == null && Age(queued) > 1500) { inputStatus = "입력 대기로 좌표 클릭 갱신 중"; return; }
                    if (frame != null)
                    {
                        IntPtr current = frame.Voice ? WindowsMacro.Instance.FindVoiceRoomWindow()
                            : WindowsMacro.Instance.FindTargetWindow(frame.Target);
                        if (frame.Session != session || current == IntPtr.Zero || current != frame.Handle
                            || WindowsMacro.Instance.GetWindowPos(current) != frame.Position
                            || WindowsMacro.Instance.GetWindowSize(current) != frame.Size)
                        { inputStatus = "대상 창 변경으로 재인식 중"; return; }
                        if (Age(frame.Tick) > 1500) { inputStatus = "인식 결과 만료로 재인식 중"; return; }
                        // 수락창과 진행자 메뉴는 보이스룸과 HWND가 다른 소유 팝업일 수 있습니다.
                        // 다른 프로그램이나 별도 카카오톡 창을 허용하지 않고 소유 관계를 검사합니다.
                        IntPtr pointWindow = WindowFromPoint(point);
                        IntPtr root = GetAncestor(current, 2);
                        IntPtr pointRoot = GetAncestor(pointWindow, 2);
                        uint targetProcess, pointProcess;
                        GetWindowThreadProcessId(current, out targetProcess);
                        GetWindowThreadProcessId(pointWindow, out pointProcess);
                        var className = new System.Text.StringBuilder(64);
                        GetClassName(pointWindow, className, className.Capacity);
                        bool allowed = IsClickWindowAllowed(root, pointRoot,
                            GetAncestor(current, 3), GetAncestor(pointWindow, 3),
                            targetProcess, pointProcess, className.ToString(),
                            GetAncestor(GetForegroundWindow(), 3));
                        if (!allowed) { inputStatus = "클릭 위치를 다른 창이 가려 재인식 대기 중"; return; }
                    }
                    else if (!IsClickMacroRunning) return;
                    inputStatus = null;
                    WindowsMacro.Instance.SetCursor(point.X, point.Y);
                    if (right) WindowsMacro.Instance.ClickRight(); else WindowsMacro.Instance.ClickLeft();
                    if (presenter) Interlocked.Exchange(ref presenterOpened, Tick);
                }
                finally { Interlocked.Exchange(ref clickPending, 0); }
            });
            if (!posted) { inputStatus = "입력 큐 접수 대기 중"; Interlocked.Exchange(ref clickPending, 0); }
            return posted;
        }

        internal static bool IsClickWindowAllowed(IntPtr root, IntPtr pointRoot,
            IntPtr owner, IntPtr pointOwner, uint process, uint pointProcess, string windowClass, IntPtr foregroundOwner)
        {
            if (root == IntPtr.Zero || pointRoot == IntPtr.Zero || process == 0 || process != pointProcess) return false;
            if (root == pointRoot) return true;
            if (owner != IntPtr.Zero && owner == pointOwner) return true;
            return windowClass == "#32768" && owner != IntPtr.Zero && foregroundOwner == owner;
        }

        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, System.Text.StringBuilder name, int capacity);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point point);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);

        // UI는 작업자가 소유한 Bitmap을 직접 표시하지 않고 변경된 프레임의 복사본만 받습니다.
        internal bool TryCopyPreviews(ref long version, out Bitmap[] images)
        {
            images = null;
            if (!Monitor.TryEnter(screenGate)) return false;
            try
            {
                if (version == screenVersion) return false;
                var copies = new Bitmap[3];
                try
                {
                    copies[0] = CurrentScreen == null ? null : (Bitmap)CurrentScreen.Clone();
                    copies[1] = CurrentScreen2 == null ? null : (Bitmap)CurrentScreen2.Clone();
                    copies[2] = CurrentScreen3 == null ? null : (Bitmap)CurrentScreen3.Clone();
                }
                catch { foreach (var copy in copies) if (copy != null) copy.Dispose(); throw; }
                images = copies; version = screenVersion;
                return true;
            }
            finally { Monitor.Exit(screenGate); }
        }

        public void Start(int x, int y, long delay)
        {
            X = x;
            Y = y;
            autoClickTimer = new CustomTimer(delay);
            IsBotRunning = true;
        }

        public void Stop()
        {
            IsBotRunning = false;
        }

        private Bitmap CaptureFrame(bool voice, int region, out Frame frame)
        {
            frame = null;
            string target = TargetWindow;
            IntPtr handle = voice ? WindowsMacro.Instance.FindVoiceRoomWindow() : WindowsMacro.Instance.FindTargetWindow(target);
            if (handle == IntPtr.Zero) return null;
            Point pos = WindowsMacro.Instance.GetWindowPos(handle);
            Point size = WindowsMacro.Instance.GetWindowSize(handle);
            if (size.X <= 0 || size.Y <= 0) return null;
            Rectangle area = region == 2
                ? new Rectangle(pos.X + (size.X - 260) / 2, pos.Y + (size.Y - 120) / 2, 130, 120)
                : new Rectangle(pos.X, pos.Y, size.X, region == 3 ? 150 : modifiedY);
            var bitmap = CaptureScreen(area);
            frame = new Frame { Handle = handle, Position = pos, Size = size, Origin = area.Location,
                Tick = Tick, Session = cycleSession, Voice = voice, Target = target };
            return bitmap;
        }

        private void ProcessCaptureScreen()
        {
            var bitmap = CaptureFrame(true, 1, out frame1);
            if (CurrentScreen != null) CurrentScreen.Dispose();
            CurrentScreen = bitmap; screenVersion++;
        }
        private void ProcessCaptureScreen2()
        {
            var bitmap = CaptureFrame(true, 2, out frame2);
            if (CurrentScreen2 != null) CurrentScreen2.Dispose();
            CurrentScreen2 = bitmap; screenVersion++;
        }
        private void ProcessCaptureScreen3()
        {
            var bitmap = CaptureFrame(false, 3, out frame3);
            if (CurrentScreen3 != null) CurrentScreen3.Dispose();
            CurrentScreen3 = bitmap; screenVersion++;
        }

        private static bool Match(Bitmap screen, Bitmap template, out Point at)
        {
            return TryFindTemplate_Sampled(screen, template, out at, 15, 1, 1, 120);
        }
        private void ProcessLineDetect()
        {
            Point at;
            modifiedY = Match(CurrentScreen, line, out at) ? at.Y + line.Height : 560;
        }
        private bool MatchAndClick(Bitmap screen, Frame frame, Bitmap template, bool right = false, bool presenter = false, bool center = false)
        {
            Point at;
            if (frame == null || Volatile.Read(ref clickPending) != 0 || !Match(screen, template, out at)) return false;
            return QueueClick(frame, new Point(frame.Origin.X + at.X + (center ? template.Width / 2 : 0),
                frame.Origin.Y + at.Y + (center ? template.Height / 2 : 0)), right, presenter);
        }
        private void ProcessAutoPresenter()
        {
            if (!MatchAndClick(CurrentScreen, frame1, manager, true, true))
                MatchAndClick(CurrentScreen, frame1, manager2, true, true);
        }
        private void ProcessAutoPresenter2()
        {
            if (MatchAndClick(CurrentScreen, frame1, host)) Interlocked.Exchange(ref presenterOpened, 0);
        }
        private void ProcessAccept()
        {
            if (!MatchAndClick(CurrentScreen2, frame2, accept)
                && !MatchAndClick(CurrentScreen2, frame2, accept2)) MatchAndClick(CurrentScreen2, frame2, check);
        }
        private void ProcessJoin() { MatchAndClick(CurrentScreen3, frame3, join, center: true); }
        private void ProcessVoiceRoomExit() { MatchAndClick(CurrentScreen2, frame2, voiceRoomExit); }

        private Bitmap CaptureScreen(Rectangle rect)
        {
            Bitmap bmp = new Bitmap(rect.Width, rect.Height);
            try
            {
                using (Graphics g = Graphics.FromImage(bmp)) g.CopyFromScreen(rect.Location, Point.Empty, rect.Size);
                return bmp;
            }
            catch { bmp.Dispose(); throw; }
        }

        public readonly struct SamplePoint
        {
            public readonly int X, Y;
            public SamplePoint(int x, int y) { X = x; Y = y; }
        }

        public static SamplePoint[] BuildGridSamples(int w, int h, int gridStep = 6, int maxPoints = 120)
        {
            var pts = new List<SamplePoint>(maxPoints);

            void Add(int x, int y)
            {
                x = ClampInt(x, 0, w - 1);
                y = ClampInt(y, 0, h - 1);
                pts.Add(new SamplePoint(x, y));
            }

            // 고정 포인트(코너/중앙/사분면)
            Add(w / 2, h / 2);
            Add(w / 4, h / 4);
            Add(3 * w / 4, h / 4);
            Add(w / 4, 3 * h / 4);
            Add(3 * w / 4, 3 * h / 4);

            Add(1, 1);
            Add(w - 2, 1);
            Add(1, h - 2);
            Add(w - 2, h - 2);

            // 격자 샘플
            for (int y = 0; y < h; y += gridStep)
            {
                for (int x = 0; x < w; x += gridStep)
                {
                    pts.Add(new SamplePoint(x, y));
                    if (pts.Count >= maxPoints) return pts.ToArray();
                }
            }

            return pts.ToArray();
        }

        public static bool TryFindTemplate_Sampled(
            Bitmap source,
            Bitmap template,
            out Point foundAt,
            int tolerance = 12,
            int searchStep = 1,
            int gridSampleStep = 6,
            int maxSamplePoints = 120
        )
        {
            foundAt = default;

            if (source == null || template == null || template.Width > source.Width || template.Height > source.Height) return false;
            Bitmap src32 = null, tpl32 = null;
            BitmapData ds = null, dt = null;
            try
            {
                src32 = Ensure32bppArgb(source);
                tpl32 = Ensure32bppArgb(template);
                int sw = src32.Width, sh = src32.Height;
                int tw = tpl32.Width, th = tpl32.Height;
                var samples = BuildGridSamples(tw, th, gridSampleStep, maxSamplePoints);
                var rectS = new Rectangle(0, 0, sw, sh);
                var rectT = new Rectangle(0, 0, tw, th);
                ds = src32.LockBits(rectS, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                dt = tpl32.LockBits(rectT, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

                int sStride = ds.Stride;
                int tStride = dt.Stride;

                int sBytes = Math.Abs(sStride) * sh;
                int tBytes = Math.Abs(tStride) * th;

                byte[] s = GetComparisonBuffer(ref sourceBuffer, sBytes);
                byte[] t = GetComparisonBuffer(ref templateBuffer, tBytes);

                Marshal.Copy(ds.Scan0, s, 0, sBytes);
                Marshal.Copy(dt.Scan0, t, 0, tBytes);

                int maxX = sw - tw;
                int maxY = sh - th;

                // 앵커(중앙) 하나로 먼저 거르는 것도 추가 (샘플링 전에 1픽셀 필터)
                int ax = tw / 2, ay = th / 2;
                int tAnchor = ay * tStride + ax * 4;
                byte tB = t[tAnchor + 0], tG = t[tAnchor + 1], tR = t[tAnchor + 2];

                for (int y = 0; y <= maxY; y += searchStep)
                {
                    for (int x = 0; x <= maxX; x += searchStep)
                    {
                        int sAnchor = (y + ay) * sStride + (x + ax) * 4;
                        if (AbsDiff(s[sAnchor + 0], tB) > tolerance) continue;
                        if (AbsDiff(s[sAnchor + 1], tG) > tolerance) continue;
                        if (AbsDiff(s[sAnchor + 2], tR) > tolerance) continue;

                        // 1차: 샘플 포인트만 비교
                        if (!SamplesPass(s, t, sStride, tStride, x, y, samples, tolerance))
                            continue;

                        // 2차: 정밀검증(전체 픽셀)
                        if (FullMatchAt_Fuzzy(s, t, sStride, tStride, x, y, tw, th, tolerance, maxBadRatio: 0.03))
                        {
                            foundAt = new Point(x, y);
                            return true;
                        }
                    }
                }

                return false;
            }
            finally
            {
                if (dt != null) tpl32.UnlockBits(dt);
                if (ds != null) src32.UnlockBits(ds);

                if (src32 != null && !ReferenceEquals(src32, source)) src32.Dispose();
                if (tpl32 != null && !ReferenceEquals(tpl32, template)) tpl32.Dispose();
            }
        }

        private static bool SamplesPass(
            byte[] s,
            byte[] t,
            int sStride,
            int tStride,
            int startX,
            int startY,
            SamplePoint[] samples,
            int tol
        )
        {
            for (int i = 0; i < samples.Length; i++)
            {
                var p = samples[i];

                int iS = (startY + p.Y) * sStride + (startX + p.X) * 4;
                int iT = p.Y * tStride + p.X * 4;

                if (AbsDiff(s[iS + 0], t[iT + 0]) > tol) return false; // B
                if (AbsDiff(s[iS + 1], t[iT + 1]) > tol) return false; // G
                if (AbsDiff(s[iS + 2], t[iT + 2]) > tol) return false; // R
            }
            return true;
        }

        private static bool FullMatchAt(
            byte[] s,
            byte[] t,
            int sStride,
            int tStride,
            int startX,
            int startY,
            int tw,
            int th,
            int tol
        )
        {
            for (int ty = 0; ty < th; ty++)
            {
                int sRow = (startY + ty) * sStride + startX * 4;
                int tRow = ty * tStride;

                for (int tx = 0; tx < tw; tx++)
                {
                    int iS = sRow + tx * 4;
                    int iT = tRow + tx * 4;

                    if (AbsDiff(s[iS + 0], t[iT + 0]) > tol) return false;
                    if (AbsDiff(s[iS + 1], t[iT + 1]) > tol) return false;
                    if (AbsDiff(s[iS + 2], t[iT + 2]) > tol) return false;
                }
            }
            return true;
        }

        private static bool FullMatchAt_Fuzzy(
    byte[] s, byte[] t,
    int sStride, int tStride,
    int startX, int startY,
    int tw, int th,
    int tol,
    double maxBadRatio = 0.02,   // 전체 픽셀 중 최대 2%까지 틀려도 OK
    int perPixelTolBoost = 0     // 필요하면 5~10 올려서 더 관대하게
)
        {
            int total = tw * th;
            int badLimit = (int)(total * maxBadRatio);
            int bad = 0;
            int tol2 = tol + perPixelTolBoost;

            for (int ty = 0; ty < th; ty++)
            {
                int sRow = (startY + ty) * sStride + startX * 4;
                int tRow = ty * tStride;

                for (int tx = 0; tx < tw; tx++)
                {
                    int iS = sRow + tx * 4;
                    int iT = tRow + tx * 4;

                    // BGR 비교(알파는 무시)
                    if (AbsDiff(s[iS + 0], t[iT + 0]) > tol2 ||
                        AbsDiff(s[iS + 1], t[iT + 1]) > tol2 ||
                        AbsDiff(s[iS + 2], t[iT + 2]) > tol2)
                    {
                        bad++;
                        if (bad > badLimit) return false;
                    }
                }
            }
            return true;
        }

        private static int AbsDiff(byte a, byte b) => a > b ? a - b : b - a;

        // 한 입력 스레드에서 반복 비교할 때 대형 배열을 매번 할당하지 않습니다.
        [ThreadStatic] private static byte[] sourceBuffer;
        [ThreadStatic] private static byte[] templateBuffer;
        private static byte[] GetComparisonBuffer(ref byte[] buffer, int length)
        {
            const int limit = 16 * 1024 * 1024;
            if (length > limit) return new byte[length];
            if (buffer == null || buffer.Length < length) buffer = new byte[length];
            return buffer;
        }

        private static Bitmap Ensure32bppArgb(Bitmap src)
        {
            if (src.PixelFormat == PixelFormat.Format32bppArgb)
                return src;

            var bmp = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
            try
            {
                using (var g = Graphics.FromImage(bmp)) g.DrawImage(src, 0, 0, src.Width, src.Height);
                return bmp;
            }
            catch { bmp.Dispose(); throw; }
        }

        private static int ClampInt(int v, int min, int max)
        {
            if (v < min) return min;
            if (v > max) return max;
            return v;
        }
    }
}
