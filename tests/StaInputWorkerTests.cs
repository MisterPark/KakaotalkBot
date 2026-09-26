using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using KakaotalkBot;

class StaInputWorkerTests
{
    static void Main()
    {
        var worker = StaInputWorker.Instance;
        int caller = Thread.CurrentThread.ManagedThreadId;
        int input = worker.Invoke(() => Thread.CurrentThread.ManagedThreadId);
        Check(input != caller, "호출 스레드와 분리");
        Check(worker.Invoke(() => Thread.CurrentThread.GetApartmentState()) == ApartmentState.STA, "STA 확인");
        Check(worker.Invoke(() => worker.Invoke(() => Thread.CurrentThread.ManagedThreadId)) == input, "중첩 호출 교착 없음");
        var entries = new List<int>();
        var tasks = new List<Task>();
        for (int n = 0; n < 24; n++)
        {
            int id = n;
            tasks.Add(Task.Run(() => worker.Invoke(() =>
            {
                if (Thread.CurrentThread.ManagedThreadId != input) throw new Exception("입력 스레드 변경");
                entries.Add(id);
                Thread.Sleep(2);
                worker.Invoke(() => entries.Add(id));
            })));
        }
        Check(Task.WaitAll(tasks.ToArray(), 10000), "동시 호출 완료");
        for (int n = 0; n < entries.Count; n += 2)
            if (entries[n] != entries[n + 1]) throw new Exception("입력 묶음이 섞임");
        Check(entries.Count == 48, "입력 묶음 직렬 실행 및 누락 없음");
        bool failed = false;
        try { worker.Invoke(() => { throw new InvalidOperationException("시험 오류"); }); }
        catch (InvalidOperationException e) { failed = e.Message == "시험 오류"; }
        Check(failed, "원래 예외 전달");
        Check(worker.Invoke(() => 42) == 42, "오류 이후 후속 작업 실행");
        var limited = new StaInputWorker(2);
        var entered = new ManualResetEventSlim(false);
        var release = new ManualResetEventSlim(false);
        int executed = 0;
        var held = Task.Factory.StartNew(() => limited.Invoke(() => { entered.Set(); release.Wait(); Interlocked.Increment(ref executed); }), TaskCreationOptions.LongRunning);
        Check(entered.Wait(5000), "포화 시험 작업 시작");
        var first = Task.Factory.StartNew(() => limited.Invoke(() => Interlocked.Increment(ref executed)), TaskCreationOptions.LongRunning);
        var second = Task.Factory.StartNew(() => limited.Invoke(() => Interlocked.Increment(ref executed)), TaskCreationOptions.LongRunning);
        try
        {
            Check(SpinWait.SpinUntil(() => limited.PendingCount == 2, 5000), "입력 큐 용량 제한 유지");
            var waiting = new ManualResetEventSlim(false);
            var overflow = Task.Factory.StartNew(() => { waiting.Set(); limited.Invoke(() => Interlocked.Increment(ref executed)); }, TaskCreationOptions.LongRunning);
            Check(waiting.Wait(5000), "초과 호출 시작");
            Check(!overflow.Wait(150), "큐 포화 시 예외 대신 대기");
            release.Set();
            Check(Task.WaitAll(new[] { held, first, second, overflow }, 5000), "포화 해소 후 모든 호출 완료");
            Check(executed == 4, "포화 중 작업 누락·중복 없음");
        }
        finally { release.Set(); limited.Complete(); }
        worker.Complete();
        bool closed = false;
        try { worker.Invoke(() => { }); }
        catch (InvalidOperationException) { closed = true; }
        Check(closed, "종료 이후 접수 차단");
    }
    static void Check(bool ok, string name)
    {
        if (!ok) throw new Exception(name);
        Console.WriteLine("PASS " + name);
    }
}
