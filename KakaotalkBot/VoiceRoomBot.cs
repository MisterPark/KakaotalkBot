using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace KakaotalkBot
{
    public class VoiceRoomBot
    {
        private bool isBotRunning = false;
        private CustomTimer autoClickTimer;
        private CustomTimer screenCaptureTimer = new CustomTimer(100);
        private CustomTimer screenCaptureTimer2 = new CustomTimer(90);
        private CustomTimer autoPresenterTimer = new CustomTimer(10000);
        private CustomTimer autoPresenterTimer2 = new CustomTimer(1000);
        private CustomTimer lineDetectorTimer = new CustomTimer(10000);
        private CustomTimer acceptTimer = new CustomTimer(90);

        private int modifiedY = 414;
        private bool isLineFound = false;
        private Bitmap line;
        private Bitmap manager;
        private Bitmap manager2;
        private Bitmap host;
        private Bitmap accept;

        public int X { get; set; } = 0;
        public int Y { get; set; } = 0;
        public bool IsClickMacroRunning { get; set; } = false;
        public bool IsBotRunning
        {
            get { return isBotRunning; }
            set
            {
                isBotRunning = value;
            }
        }
        public Bitmap CurrentScreen { get; private set; }
        public Bitmap CurrentScreen2 { get; private set; }


        public VoiceRoomBot()
        {
            line = new Bitmap("리스너경계선.bmp");
            manager = new Bitmap("방장.bmp");
            manager2 = new Bitmap("부방장.bmp");
            host = new Bitmap("진행자.bmp");
            accept = new Bitmap("수락.bmp");
        }

        public void Update()
        {
            if (IsBotRunning == false) return;

            if (autoClickTimer.Check(Time.DeltaTime))
            {
                if (IsClickMacroRunning)
                {
                    WindowsMacro.Instance.SetCursor(X, Y);
                    WindowsMacro.Instance.ClickLeft();
                }
            }

            if (screenCaptureTimer.Check(Time.DeltaTime))
            {
                ProcessCaptureScreen();
            }

            if (screenCaptureTimer2.Check(Time.DeltaTime))
            {
                ProcessCaptureScreen2();
            }

            if (acceptTimer.Check(Time.DeltaTime))
            {
                ProcessAccept();
            }

            if (lineDetectorTimer.Check(Time.DeltaTime))
            {
                ProcessLineDetect();
            }

            if (autoPresenterTimer.Check(Time.DeltaTime))
            {
                ProcessAutoPresenter();
            }

            if (autoPresenterTimer2.Check(Time.DeltaTime))
            {
                ProcessAutoPresenter2();
            }

        }

        public void Start(int x, int y, long delay)
        {
            X = x;
            Y = y;
            autoClickTimer = new CustomTimer(delay);
            isBotRunning = true;
        }

        public void Stop()
        {
            isBotRunning = false;
        }

        private void ProcessCaptureScreen()
        {
            IntPtr handle = WindowsMacro.Instance.FindVoiceRoomWindow();
            if (handle == IntPtr.Zero) return;

            Point pos = WindowsMacro.Instance.GetWindowPos(handle);
            Point size = WindowsMacro.Instance.GetWindowSize(handle);

            Rectangle captureArea = new Rectangle(pos.X, pos.Y, size.X, modifiedY);


            if (CurrentScreen != null)
            {
                CurrentScreen.Dispose();
                CurrentScreen = null;
            }
            CurrentScreen = CaptureScreen(captureArea);
        }

        private void ProcessCaptureScreen2()
        {
            IntPtr handle = WindowsMacro.Instance.FindVoiceRoomWindow();
            if (handle == IntPtr.Zero) return;

            Point pos = WindowsMacro.Instance.GetWindowPos(handle);
            Point size = WindowsMacro.Instance.GetWindowSize(handle);

            int w = 260;
            int h = 120;
            int x = pos.X + (size.X - w) / 2;
            int y = pos.Y + (size.Y - h) / 2;

            Rectangle captureArea = new Rectangle(x, y, w, h);


            if (CurrentScreen2 != null)
            {
                CurrentScreen2.Dispose();
                CurrentScreen2 = null;
            }
            CurrentScreen2 = CaptureScreen(captureArea);
        }

        private void ProcessLineDetect()
        {
            IntPtr handle = WindowsMacro.Instance.FindVoiceRoomWindow();
            if (handle == IntPtr.Zero) return;

            if (TryFindTemplate_Sampled(
                    CurrentScreen, line, out var at2,
                    tolerance: 15,
                    searchStep: 1,
                    gridSampleStep: 1,
                    maxSamplePoints: 120))
            {
                Point pos = WindowsMacro.Instance.GetWindowPos(handle);
                int x = at2.X + pos.X;
                int y = at2.Y + pos.Y;

                modifiedY = at2.Y + line.Height;
            }
            else
            {
                modifiedY = 414;
            }

        }

        private void ProcessAutoPresenter()
        {
            IntPtr handle = WindowsMacro.Instance.FindVoiceRoomWindow();
            if (handle == IntPtr.Zero) return;

            if (TryFindTemplate_Sampled(
                       CurrentScreen, manager, out var at,
                       tolerance: 15,
                       searchStep: 1,
                       gridSampleStep: 1,
                       maxSamplePoints: 120))
            {
                Point pos = WindowsMacro.Instance.GetWindowPos(handle);
                int x = at.X + pos.X;
                int y = at.Y + pos.Y;
                WindowsMacro.Instance.SetCursor(x, y);
                WindowsMacro.Instance.ClickRight();
            }

            if (TryFindTemplate_Sampled(
                     CurrentScreen, manager2, out var at2,
                     tolerance: 15,
                     searchStep: 1,
                     gridSampleStep: 1,
                     maxSamplePoints: 120))
            {
                Point pos = WindowsMacro.Instance.GetWindowPos(handle);
                int x = at2.X + pos.X;
                int y = at2.Y + pos.Y;
                WindowsMacro.Instance.SetCursor(x, y);
                WindowsMacro.Instance.ClickRight();
            }
        }

        private void ProcessAutoPresenter2()
        {
            IntPtr handle = WindowsMacro.Instance.FindVoiceRoomWindow();
            if (handle == IntPtr.Zero) return;

            if (TryFindTemplate_Sampled(
                     CurrentScreen, host, out var at3,
                     tolerance: 15,
                     searchStep: 1,
                     gridSampleStep: 1,
                     maxSamplePoints: 120))
            {
                Point pos = WindowsMacro.Instance.GetWindowPos(handle);
                int x2 = at3.X + pos.X;
                int y2 = at3.Y + pos.Y;
                WindowsMacro.Instance.SetCursor(x2, y2);
                WindowsMacro.Instance.ClickLeft();
            }
        }

        private void ProcessAccept()
        {
            IntPtr handle = WindowsMacro.Instance.FindVoiceRoomWindow();
            if (handle == IntPtr.Zero) return;

            Point pos = WindowsMacro.Instance.GetWindowPos(handle);
            Point size = WindowsMacro.Instance.GetWindowSize(handle);

            int w = 260;
            int h = 120;
            int x = pos.X + (size.X - w) / 2;
            int y = pos.Y + (size.Y - h) / 2;

            if (TryFindTemplate_Sampled(
                    CurrentScreen2, accept, out var at,
                    tolerance: 15,
                    searchStep: 1,
                    gridSampleStep: 1,
                    maxSamplePoints: 120))
            {
                int x2 = x + at.X;
                int y2 = y + at.Y;
                WindowsMacro.Instance.SetCursor(x2, y2);
                WindowsMacro.Instance.ClickLeft();
            }
        }

        private Bitmap CaptureScreen(Rectangle rect)
        {
            Bitmap bmp = new Bitmap(rect.Width, rect.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(rect.Location, Point.Empty, rect.Size);
            }
            return bmp;
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

            Bitmap src32 = Ensure32bppArgb(source);
            Bitmap tpl32 = Ensure32bppArgb(template);

            int sw = src32.Width, sh = src32.Height;
            int tw = tpl32.Width, th = tpl32.Height;

            if (tw > sw || th > sh) return false;

            var samples = BuildGridSamples(tw, th, gridSampleStep, maxSamplePoints);

            var rectS = new Rectangle(0, 0, sw, sh);
            var rectT = new Rectangle(0, 0, tw, th);

            BitmapData ds = null;
            BitmapData dt = null;

            try
            {
                ds = src32.LockBits(rectS, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                dt = tpl32.LockBits(rectT, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

                int sStride = ds.Stride;
                int tStride = dt.Stride;

                int sBytes = Math.Abs(sStride) * sh;
                int tBytes = Math.Abs(tStride) * th;

                byte[] s = new byte[sBytes];
                byte[] t = new byte[tBytes];

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

                src32.Dispose();
                src32 = null;
                tpl32.Dispose();
                tpl32 = null;
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

        private static Bitmap Ensure32bppArgb(Bitmap src)
        {
            if (src.PixelFormat == PixelFormat.Format32bppArgb)
                return (Bitmap)src.Clone();

            var bmp = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
                g.DrawImage(src, 0, 0, src.Width, src.Height);
            return bmp;
        }

        private static int ClampInt(int v, int min, int max)
        {
            if (v < min) return min;
            if (v > max) return max;
            return v;
        }
    }
}
