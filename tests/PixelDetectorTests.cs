using System;
using System.Drawing;
using System.Threading;
using KakaotalkBot;
class PixelDetectorTests
{
    static void Check(bool ok, string text) { if (!ok) throw new Exception(text); }
    static int Main()
    {
        ScreenPixelDetector detector = null;
        try
        {
            int reads = 0, calls = 0;
            detector = new ScreenPixelDetector((x, y) => Interlocked.Increment(ref reads) % 2 == 0 ? Color.Red : Color.Blue);
            Action listener = () => Interlocked.Increment(ref calls);
            detector.AddListener(listener); detector.AddListener(listener);
            detector.Invoke(); Check(calls == 1, "같은 핸들러 중복 등록");
            detector.Start(1, 1);
            long first = detector.Generation;
            Check(SpinWait.SpinUntil(() => Volatile.Read(ref reads) >= 4, 3000), "픽셀 주기 실행");
            detector.Stop();
            Check(!detector.IsCurrentSession(first), "중지된 입력 무효화");
            detector.Start(2, 2);
            Check(detector.Generation != first && detector.IsRunning, "중지 후 새 감지 스레드 시작");
            int before = reads;
            Thread.Sleep(260);
            Check(reads - before <= 8, "무한 폴링 방지");
            detector.Stop(); detector.RemoveListener(listener);
            int previous = calls; detector.Invoke(); Check(calls == previous, "이전 핸들러 제거");
            detector.AddListener(() => { throw new InvalidOperationException("시험 오류"); });
            detector.Start(3, 3);
            Check(SpinWait.SpinUntil(() => !detector.IsRunning, 3000), "콜백 오류 격리");
            Check(detector.LastError != null, "콜백 오류 상태 기록");
            Console.WriteLine("PASS: polling bounds, listener deduplication, stop/restart, stale session rejection, callback failure isolation (no native input)");
            return 0;
        }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
        finally { if (detector != null) detector.Stop(); }
    }
}
