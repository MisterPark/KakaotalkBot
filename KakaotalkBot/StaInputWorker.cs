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
            var result = new TaskCompletionSource<T>();
            Action work = () =>
            {
                try { result.SetResult(action()); }
                catch (Exception error) { result.SetException(error); }
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
}
