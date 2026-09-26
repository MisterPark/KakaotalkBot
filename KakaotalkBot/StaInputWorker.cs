using System;
using System.Collections.Generic;
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
        private readonly object queueGate = new object();
        private readonly Queue<Action> chatQueue = new Queue<Action>();
        private readonly Queue<Action> backgroundQueue = new Queue<Action>();
        private readonly HashSet<object> backgroundKeys = new HashSet<object>();
        private readonly AutoResetEvent ready = new AutoResetEvent(false);
        private readonly int capacity;
        private bool completed;
        private volatile string backgroundError;
        internal string BackgroundError { get { return backgroundError; } }
        private readonly Thread thread;

        private StaInputWorker() : this(256) { }

        internal StaInputWorker(int capacity)
        {
            if (capacity <= 0) throw new ArgumentOutOfRangeException("capacity");
            this.capacity = capacity;
            thread = new Thread(Run) { IsBackground = true, Name = "카카오톡 입력 STA" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        internal int PendingCount { get { lock (queueGate) return chatQueue.Count + backgroundQueue.Count; } }

        internal bool IsCurrent { get { return Thread.CurrentThread == thread; } }

        internal void Invoke(Action action)
        {
            Invoke<object>(() => { action(); return null; });
        }

        internal T Invoke<T>(Func<T> action)
        {
            // 같은 작업 안의 입력 함수 호출은 큐에 다시 넣지 않습니다.
            if (IsCurrent) return action();
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
            // 큐 포화는 정상적인 대기 상태입니다. 빈자리가 생길 때까지 생산자를 기다리게 합니다.
            // 무제한 적재·작업 누락 없이 기존 입력 순서를 유지합니다. 종료 후 접수는 기존대로 거부합니다.
            lock (queueGate)
            {
                while (!completed && chatQueue.Count >= capacity) Monitor.Wait(queueGate);
                if (completed) throw new InvalidOperationException("입력 작업자가 종료되었습니다.");
                chatQueue.Enqueue(work);
                ready.Set();
            }
            return result.Task.GetAwaiter().GetResult();
        }

        // 창 유지보수의 예상 가능한 실패는 작업자 안에서 결과로 반환합니다.
        // 일반 전송 Invoke의 오류 전달 동작은 유지합니다.
        internal bool TryInvoke<T>(Func<T> action, out T value, out string error)
        {
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
            lock (queueGate)
            {
                if (completed || backgroundKeys.Contains(key) || backgroundQueue.Count >= capacity) return false;
                backgroundKeys.Add(key);
                long queuedAt = System.Diagnostics.Stopwatch.GetTimestamp();
                backgroundQueue.Enqueue(() =>
                {
                    long started = System.Diagnostics.Stopwatch.GetTimestamp();
                    DelayDiagnostics.Record("background-queue:" + action.Method.Name, queuedAt);
                    try { action(); backgroundError = null; }
                    catch (Exception error) { backgroundError = error.GetType().Name + ": " + error.Message; }
                    finally
                    {
                        DelayDiagnostics.Record("background-action:" + action.Method.Name, started);
                        lock (queueGate) backgroundKeys.Remove(key);
                    }
                });
                ready.Set();
                return true;
            }
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
                    Action work = null;
                    lock (queueGate)
                    {
                        if (chatQueue.Count > 0)
                        {
                            work = chatQueue.Dequeue();
                            Monitor.PulseAll(queueGate);
                        }
                        else if (backgroundQueue.Count > 0) work = backgroundQueue.Dequeue();
                        else if (completed) return;
                    }
                    if (work != null) work();
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
            lock (queueGate)
            {
                if (completed) return;
                completed = true;
                Monitor.PulseAll(queueGate);
                ready.Set();
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
