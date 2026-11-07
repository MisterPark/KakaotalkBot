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
        bool recycleResult;
        string recycleError;
        Check(!worker.TryInvoke<bool>(() => { throw new TimeoutException("시험 재열기 시간 초과"); }, out recycleResult, out recycleError) &&
            !recycleResult && recycleError.Contains("시간 초과"), "유지보수 시간 초과를 작업자 안에서 처리");
        Check(!worker.TryInvoke<bool>(() => { throw new InvalidOperationException("시험 창 상태 변경"); }, out recycleResult, out recycleError) &&
            recycleError.Contains("창 상태"), "창 상태 실패를 결과로 반환");
        Check(worker.Invoke(() => 43) == 43, "재열기 실패 후 입력 작업자 유지");
        var macro = new StaInputWorker(8);
        var trace = new List<string>();
        var macroStarted = new ManualResetEventSlim(false);
        var macroRelease = new ManualResetEventSlim(false);
        var backgroundDone = new ManualResetEventSlim(false);
        int producer = Thread.CurrentThread.ManagedThreadId;
        Check(macro.TryBuildCommand(() =>
        {
            Check(Thread.CurrentThread.ManagedThreadId == producer, "생성은 생산자 스레드");
            macro.Invoke(() => { trace.Add("first"); macroStarted.Set(); macroRelease.Wait(); });
            macro.Invoke(() => trace.Add("last"));
        }), "매크로 명령 큐 등록");
        try
        {
            Check(macroStarted.Wait(5000), "실행 완료 전 생산자 반환");
            Check(macro.TryBuildCommand(() => macro.Invoke(() => trace.Add("second"))), "첫 명령 실행 중 다음 명령 생성·등록 완료");
            macro.TryPostBackground(new object(), () => { trace.Add("background"); backgroundDone.Set(); });
            Check(!backgroundDone.Wait(100), "명령 실행 중 보이스룸 명령 대기");
        }
        finally { macroRelease.Set(); }
        Check(backgroundDone.Wait(5000), "명령 전체 종료 후 보이스룸 실행");
        Check(string.Join(",", trace) == "first,last,second,background", "명령 내 단계 순차 실행");
        var recovered = new ManualResetEventSlim(false);
        macro.TryBuildCommand(() => macro.Invoke(() => { throw new Exception("명령 오류"); }));
        macro.TryPostBackground(new object(), () => recovered.Set());
        Check(recovered.Wait(5000), "명령 실행 예외 후 다음 명령 진행");
        Check(macro.CommandError != null, "비동기 매크로 오류 표시");
        macro.Complete();
        var closing = new StaInputWorker(4);
        var building = new ManualResetEventSlim(false);
        var finishBuild = new ManualResetEventSlim(false);
        var drained = new ManualResetEventSlim(false);
        var accepted = Task.Run(() => closing.TryBuildCommand(() =>
        {
            building.Set(); finishBuild.Wait();
            closing.Invoke(() => drained.Set());
        }));
        try
        {
            Check(building.Wait(5000), "종료 경합 시험: 명령 생성 중");
            closing.Complete();
            Check(!closing.TryBuildCommand(() => { throw new Exception("종료 후 생성됨"); }), "종료 후 새 생산자 접수 거부");
        }
        finally { finishBuild.Set(); }
        Check(accepted.Wait(5000) && accepted.Result && drained.Wait(5000), "종료와 경합한 기존 접수 명령 누락 없이 실행");

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
        var priority = new StaInputWorker(4);
        var started = new ManualResetEventSlim(false);
        var finish = new ManualResetEventSlim(false);
        var lowDone = new ManualResetEventSlim(false);
        var order = new List<string>();
        var blocker = Task.Factory.StartNew(() => priority.Invoke(() => { started.Set(); finish.Wait(); order.Add("running"); }), TaskCreationOptions.LongRunning);
        Check(started.Wait(5000), "우선순위 시험 시작");
        object lowKey = new object();
        Check(priority.TryPostBackground(lowKey, () => { order.Add("voice"); lowDone.Set(); }), "보이스룸 비동기 접수");
        Check(!priority.TryPostBackground(lowKey, () => order.Add("duplicate")), "보이스룸 중복 적재 방지");
        var chat = Task.Factory.StartNew(() => priority.Invoke(() => order.Add("chat")), TaskCreationOptions.LongRunning);
        try
        {
            Check(SpinWait.SpinUntil(() => priority.PendingCount == 2, 5000), "두 우선순위 대기 확인");
            finish.Set();
            Check(Task.WaitAll(new[] { blocker, chat }, 5000) && lowDone.Wait(5000), "우선순위 작업 완료");
            Check(string.Join(",", order) == "running,chat,voice", "실행 중 묶음 유지·채팅 우선·보이스룸 후순위");
            priority.TryPostBackground(new object(), () => { throw new InvalidOperationException("보조 입력 시험"); });
            Check(SpinWait.SpinUntil(() => priority.BackgroundError != null, 5000), "보조 입력 오류 상태 기록");
            Check(priority.Invoke(() => 99) == 99, "보조 입력 실패 후 채팅 처리 유지");
        }
        finally { finish.Set(); priority.Complete(); }
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
