using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KakaotalkBot
{
    // 입력 묶음 하나가 끝나기 전에는 다른 입력 묶음을 실행하지 않습니다.
    internal sealed class StaInputWorker
    {
        internal static readonly StaInputWorker Instance = new StaInputWorker();
        private readonly BlockingCollection<Action> queue;
        private readonly Thread thread;

        private StaInputWorker() : this(256) { }

        internal StaInputWorker(int capacity)
        {
            queue = new BlockingCollection<Action>(capacity);
            thread = new Thread(Run) { IsBackground = true, Name = "카카오톡 입력 STA" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        internal int PendingCount { get { return queue.Count; } }

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
            queue.Add(work);
            return result.Task.GetAwaiter().GetResult();
        }

        private void Run()
        {
            while (!queue.IsCompleted)
            {
                Action work;
                if (queue.TryTake(out work, 20)) work();
                // 클립보드와 COM에 필요한 메시지를 처리하되 입력 큐는 여기서 재진입하지 않습니다.
                Application.DoEvents();
            }
        }

        internal void Complete()
        {
            // 이미 접수한 작업은 완료하고 새 작업은 받지 않습니다.
            queue.CompleteAdding();
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
