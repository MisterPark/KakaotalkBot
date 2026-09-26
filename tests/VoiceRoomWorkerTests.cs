using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using KakaotalkBot;
class VoiceRoomWorkerTests
{
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); Console.WriteLine("PASS " + message); }
    static void Main()
    {
        var policy = typeof(VoiceRoomBot).GetMethod("IsClickWindowAllowed", BindingFlags.Static | BindingFlags.NonPublic);
        Func<int, int, int, int, uint, uint, string, int, bool> allowed = (root, hit, owner, hitOwner, pid, hitPid, cls, foreground) =>
            (bool)policy.Invoke(null, new object[] { new IntPtr(root), new IntPtr(hit), new IntPtr(owner), new IntPtr(hitOwner), pid, hitPid, cls, new IntPtr(foreground) });
        Check(allowed(1, 1, 1, 1, 10, 10, "", 1), "보이스룸 본체 클릭 허용");
        Check(allowed(1, 2, 1, 1, 10, 10, "", 1), "별도 HWND인 소유 수락창 클릭 허용");
        Check(allowed(1, 2, 1, 2, 10, 10, "#32768", 1), "보이스룸의 전경 컨텍스트 메뉴 허용");
        Check(!allowed(1, 2, 1, 2, 10, 20, "#32768", 1), "다른 프로그램의 메뉴 차단");
        Check(!allowed(1, 2, 1, 2, 10, 10, "", 2), "별도 채팅창이 가린 좌표 차단");
        Check(!allowed(1, 2, 1, 2, 10, 10, "#32768", 2), "다른 방의 전경 메뉴 차단");
        Check(!allowed(1, 0, 1, 0, 10, 0, "", 1), "유효한 창 없는 좌표 차단");
        var entered = new ManualResetEventSlim();
        var release = new ManualResetEventSlim();
        int calls = 0, recognitionThread = 0;
        Action recognize = () => {
            recognitionThread = Thread.CurrentThread.ManagedThreadId;
            Interlocked.Increment(ref calls);
            entered.Set();
            if (!release.Wait(5000)) throw new Exception("인식 대기 시간 초과");
        };
        var bot = (VoiceRoomBot)Activator.CreateInstance(typeof(VoiceRoomBot), BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { recognize }, null);
        var staType = typeof(VoiceRoomBot).Assembly.GetType("KakaotalkBot.StaInputWorker");
        var sta = staType.GetField("Instance", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
        MethodInfo invoke = null;
        foreach (var method in staType.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic))
            if (method.Name == "Invoke" && !method.IsGenericMethod) invoke = method;
        bot.IsBotRunning = true;
        Check(entered.Wait(3000), "인식 전용 작업자 시작");
        int inputThread = 0;
        var input = Task.Run(() => invoke.Invoke(sta, new object[] { (Action)(() => { inputThread = Thread.CurrentThread.ManagedThreadId; }) }));
        Check(input.Wait(2000), "인식이 대기 중이어도 STA 입력 진행");
        Check(inputThread != recognitionThread && recognitionThread != Thread.CurrentThread.ManagedThreadId, "인식·입력·호출 스레드 분리");
        var worker = (Thread)typeof(VoiceRoomBot).GetField("recognitionThread", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(bot);
        bot.Stop(); bot.IsBotRunning = true; bot.Stop();
        Check(calls == 1, "재시작 중 인식 작업 중복 실행 없음");
        release.Set();
        Thread.Sleep(250);
        Check(calls == 1, "중지 후 인식 반복 없음");
        bot.IsBotRunning = true;
        Check(SpinWait.SpinUntil(() => Volatile.Read(ref calls) > 1, 2000), "재시작 후 인식 재개");
        Check(object.ReferenceEquals(worker, typeof(VoiceRoomBot).GetField("recognitionThread", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(bot)), "동일 작업자 재사용");
        bot.Dispose(); bot.Dispose();
        Check(worker.Join(3000), "종료 시 인식 작업자 종료");
        staType.GetMethod("Complete", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(sta, null);
    }
}
