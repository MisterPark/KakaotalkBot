using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KakaotalkBot
{
    // 입력 묶음 하나가 끝나기 전에는 다른 입력 묶음을 실행하지 않습니다.
    internal sealed class StaInputWorker
    {
        internal static readonly StaInputWorker Instance = new StaInputWorker();
        [ThreadStatic] private static List<Action> buildingSteps;
        [ThreadStatic] private static StaInputWorker buildingWorker;
        // 최상위 비트는 접수 종료, 나머지는 등록 중인 생산자 수입니다.
        private int admissionState;
        private readonly CancellationTokenSource shutdown = new CancellationTokenSource();
        private readonly SemaphoreSlim chatSlots, backgroundSlots;
        private volatile string commandError;
        internal string CommandError { get { return commandError; } }

        private sealed class MacroCommand
        {
            internal readonly Action[] Steps;
            internal MacroCommand(params Action[] steps) { Steps = (Action[])steps.Clone(); }
            internal void Execute() { foreach (var step in Steps) step(); }
        }
        private readonly ConcurrentQueue<MacroCommand> chatQueue = new ConcurrentQueue<MacroCommand>();
        private readonly ConcurrentQueue<MacroCommand> backgroundQueue = new ConcurrentQueue<MacroCommand>();
        private readonly ConcurrentDictionary<object, byte> backgroundKeys = new ConcurrentDictionary<object, byte>();
        private readonly AutoResetEvent ready = new AutoResetEvent(false);
        private volatile string backgroundError;
        internal string BackgroundError { get { return backgroundError; } }
        private readonly Thread thread;

        private StaInputWorker() : this(256) { }

        internal StaInputWorker(int capacity)
        {
            if (capacity <= 0) throw new ArgumentOutOfRangeException("capacity");
            chatSlots = new SemaphoreSlim(capacity, capacity);
            backgroundSlots = new SemaphoreSlim(capacity, capacity);
            thread = new Thread(Run) { IsBackground = true, Name = "카카오톡 입력 STA" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        internal int PendingCount { get { return chatQueue.Count + backgroundQueue.Count; } }

        internal bool IsCurrent { get { return Thread.CurrentThread == thread; } }

        internal void Invoke(Action action)
        {
            if (!IsCurrent && buildingWorker == this) { buildingSteps.Add(action); return; }
            Invoke<object>(() => { action(); return null; });
        }

        internal T Invoke<T>(Func<T> action)
        {
            // 같은 작업 안의 입력 함수 호출은 큐에 다시 넣지 않습니다.
            if (IsCurrent) return action();
            if (buildingWorker == this) throw new InvalidOperationException("명령 생성 중에는 입력 실행 결과를 기다릴 수 없습니다.");
            return InvokeQueued(action);
        }

        private T InvokeQueued<T>(Func<T> action)
        {
            long queuedAt = System.Diagnostics.Stopwatch.GetTimestamp();
            var result = new TaskCompletionSource<T>();
            Action work = () =>
            {
                long started = System.Diagnostics.Stopwatch.GetTimestamp();
                DelayDiagnostics.Record("input-queue:" + action.Method.Name, queuedAt);
                try { result.SetResult(action()); }
                catch (Exception error) { result.SetException(error); }
                finally { DelayDiagnostics.Record("input-action:" + action.Method.Name, started); }
            };
            if (!EnterProducer()) throw new InvalidOperationException("입력 작업자가 종료되었습니다.");
            try
            {
                try { chatSlots.Wait(shutdown.Token); }
                catch (OperationCanceledException) { throw new InvalidOperationException("입력 작업자가 종료되었습니다."); }
                chatQueue.Enqueue(new MacroCommand(work));
                Signal();
            }
            finally { ExitProducer(); }
            return result.Task.GetAwaiter().GetResult();
        }

        // 생산자는 매크로 단계를 모아 불변 명령을 등록하고, 실행 완료를 기다리지 않습니다.
        // 큐가 가득 차면 생성 함수를 호출하지 않아 원본 명령을 다음 주기에 처리할 수 있습니다.
        internal bool TryBuildCommand(Action build, Func<bool> isCurrent = null)
        {
            if (buildingWorker != null) throw new InvalidOperationException("명령 생성은 중첩할 수 없습니다.");
            if (!EnterProducer()) return false;
            if (!chatSlots.Wait(0)) { ExitProducer(); return false; }
            bool enqueued = false;
            var steps = new List<Action>();
            try
            {
                buildingWorker = this; buildingSteps = steps;
                try { build(); }
                finally { buildingWorker = null; buildingSteps = null; }
                var command = new MacroCommand(steps.ToArray());
                if (command.Steps.Length != 0)
                {
                    chatQueue.Enqueue(new MacroCommand(() =>
                    {
                        try
                        {
                            if (isCurrent == null || isCurrent()) command.Execute();
                            commandError = null;
                        }
                        catch (Exception error) { commandError = "매크로 실행 오류: " + error.GetBaseException().Message; }
                    }));
                    enqueued = true;
                    Signal();
                }
                return true;
            }
            finally
            {
                if (!enqueued) chatSlots.Release();
                ExitProducer();
            }
        }

        // 창 유지보수의 예상 가능한 실패는 작업자 안에서 결과로 반환합니다.
        // 일반 전송 Invoke의 오류 전달 동작은 유지합니다.
        internal bool TryInvoke<T>(Func<T> action, out T value, out string error)
        {
            if (buildingWorker == this) throw new InvalidOperationException("명령 생성 중에는 동기 입력 결과를 요청할 수 없습니다.");
            T localValue = default(T);
            string localError = null;
            Invoke(() =>
            {
                try { localValue = action(); }
                catch (Exception failure)
                {
                    if (!(failure is InvalidOperationException || failure is TimeoutException ||
                        failure is System.ComponentModel.Win32Exception || failure is System.IO.IOException ||
                        failure is System.Runtime.InteropServices.COMException)) throw;
                    localError = failure.GetType().Name + ": " + failure.Message;
                }
            });
            value = localValue; error = localError;
            return localError == null;
        }

        // 보이스룸 등 반복 작업은 키마다 대기·실행 중인 한 건만 허용합니다.
        // 호출 스레드는 완료를 기다리지 않으므로 다음 채팅 명령을 계속 처리할 수 있습니다.
        internal bool TryPostBackground(object key, Action action)
        {
            if (key == null || action == null) throw new ArgumentNullException();
            if (!EnterProducer()) return false;
            bool slot = false, added = false, enqueued = false;
            try
            {
                if (!backgroundSlots.Wait(0)) return false;
                slot = true;
                if (!backgroundKeys.TryAdd(key, 0)) return false;
                added = true;
                long queuedAt = System.Diagnostics.Stopwatch.GetTimestamp();
                backgroundQueue.Enqueue(new MacroCommand(() =>
                {
                    long started = System.Diagnostics.Stopwatch.GetTimestamp();
                    DelayDiagnostics.Record("background-queue:" + action.Method.Name, queuedAt);
                    try { action(); backgroundError = null; }
                    catch (Exception error) { backgroundError = error.GetType().Name + ": " + error.Message; }
                    finally
                    {
                        DelayDiagnostics.Record("background-action:" + action.Method.Name, started);
                        byte ignored; backgroundKeys.TryRemove(key, out ignored);
                    }
                }));
                enqueued = true;
                Signal();
                return true;
            }
            finally
            {
                if (!enqueued)
                {
                    if (added) { byte ignored; backgroundKeys.TryRemove(key, out ignored); }
                    if (slot) backgroundSlots.Release();
                }
                ExitProducer();
            }
        }

        // 종료와 등록이 경합해도 접수한 명령을 작업자가 남기고 종료하지 않습니다.
        private bool EnterProducer()
        {
            while (true)
            {
                int state = Volatile.Read(ref admissionState);
                if (state < 0) return false;
                if (Interlocked.CompareExchange(ref admissionState, state + 1, state) == state) return true;
            }
        }
        private void ExitProducer() { Interlocked.Decrement(ref admissionState); Signal(); }
        private void Signal()
        {
            try { ready.Set(); }
            catch (ObjectDisposedException) { } // 마지막 생산자 반환 직후 작업자가 종료된 경우
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr[] handles, uint milliseconds, uint wakeMask, uint flags);

        private void Run()
        {
            var handles = new[] { ready.SafeWaitHandle.DangerousGetHandle() };
            try
            {
                while (true)
                {
                    MacroCommand work = null;
                    if (chatQueue.TryDequeue(out work)) chatSlots.Release();
                    else if (backgroundQueue.TryDequeue(out work)) backgroundSlots.Release();
                    else if (Volatile.Read(ref admissionState) == int.MinValue && chatQueue.IsEmpty && backgroundQueue.IsEmpty) return;
                    if (work != null) work.Execute();
                    else
                    {
                        // 고정 주기 폴링 없이 새 작업 또는 STA 윈도우 메시지 도착 시 즉시 깨어납니다.
                        uint result = MsgWaitForMultipleObjectsEx(1, handles, uint.MaxValue, 0x04FF, 0x0004);
                        if (result == uint.MaxValue) ready.WaitOne(20);
                    }
                    Application.DoEvents();
                }
            }
            finally { ready.Dispose(); }
        }

        internal void Complete()
        {
            while (true)
            {
                int state = Volatile.Read(ref admissionState);
                if (state < 0) return;
                if (Interlocked.CompareExchange(ref admissionState, state | int.MinValue, state) != state) continue;
                shutdown.Cancel();
                Signal();
                return;
            }
        }
    }

    // 개인정보 없이 느린 작업의 종류와 소요 시간만 남깁니다.
    internal static class DelayDiagnostics
    {
        private static readonly object Gate = new object();
        private static long nextHealth;
        internal static void RecordHealth(int commands, int answers)
        {
            long now = System.Diagnostics.Stopwatch.GetTimestamp();
            if (now < Interlocked.Read(ref nextHealth)) return;
            Interlocked.Exchange(ref nextHealth, now + 60 * System.Diagnostics.Stopwatch.Frequency);
            try
            {
                using (var process = System.Diagnostics.Process.GetCurrentProcess())
                {
                    string row = DateTimeOffset.Now.ToString("O") + " privateMB=" + (process.PrivateMemorySize64 / 1048576) +
                        " managedMB=" + (GC.GetTotalMemory(false) / 1048576) + " handles=" + process.HandleCount +
                        " threads=" + process.Threads.Count + " commands=" + commands + " quizAnswers=" + answers +
                        " input=" + StaInputWorker.Instance.PendingCount + " gc2=" + GC.CollectionCount(2) + "\r\n";
                    lock (Gate)
                    {
                        string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "runtime-health.log");
                        if (System.IO.File.Exists(path) && new System.IO.FileInfo(path).Length > 1024 * 1024)
                            System.IO.File.WriteAllText(path, "");
                        System.IO.File.AppendAllText(path, row);
                    }
                }
            }
            catch (System.IO.IOException) { }
            catch (UnauthorizedAccessException) { }
            catch (System.ComponentModel.Win32Exception) { }
        }

        internal static void Record(string stage, long started)
        {
            double ms = (System.Diagnostics.Stopwatch.GetTimestamp() - started) * 1000.0 / System.Diagnostics.Stopwatch.Frequency;
            if (ms < 500) return;
            try
            {
                lock (Gate)
                {
                    string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "command-delay.log");
                    if (System.IO.File.Exists(path) && new System.IO.FileInfo(path).Length > 1024 * 1024)
                        System.IO.File.WriteAllText(path, "");
                    System.IO.File.AppendAllText(path, DateTimeOffset.Now.ToString("O") + " " + stage + " " + ms.ToString("F0", System.Globalization.CultureInfo.InvariantCulture) + "ms\r\n");
                }
            }
            catch (System.IO.IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
