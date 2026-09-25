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
        private readonly BlockingCollection<Action> queue = new BlockingCollection<Action>(256);
        private readonly Thread thread;

        private StaInputWorker()
        {
            thread = new Thread(Run) { IsBackground = true, Name = "카카오톡 입력 STA" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

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
                DelayDiagnostics.Record("input-queue", queuedAt);
                try { result.SetResult(action()); }
                catch (Exception error) { result.SetException(error); }
                finally { DelayDiagnostics.Record("input-action", started); }
            };
            if (!queue.TryAdd(work)) throw new InvalidOperationException("입력 작업이 너무 많이 대기 중입니다.");
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
