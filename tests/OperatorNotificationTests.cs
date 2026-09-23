using System;
using System.Collections;
using System.IO;
using System.Reflection;
using KakaotalkBot;

class OperatorNotificationTests
{
    static Type type = typeof(Bot).Assembly.GetType("KakaotalkBot.OperatorNotificationQueue");
    static int checks;
    static object Call(object q, string method, params object[] args) { return type.GetMethod(method, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(q, args); }
    static object New(string path) { return Activator.CreateInstance(type, BindingFlags.NonPublic | BindingFlags.Instance, null, new object[] { path }, null); }
    static IList Items(object q) { return (IList)type.GetField("Items", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(q); }
    static string State(object q, int i) { return (string)Items(q)[i].GetType().GetField("State").GetValue(Items(q)[i]); }
    static void Check(bool value, string name) { if (!value) throw new Exception(name); checks++; }
    static void Add(object q, string key) { Call(q, "Enqueue", key, "account", 20L, 10L, "[테스트]\n본문", 100L); }
    static int Main(string[] args)
    {
        try
        {
            string path = Path.Combine(args[0], "outbox.json"); var q = New(path);
            Add(q, "one"); Add(q, "one"); Check(Items(q).Count == 1, "이벤트 중복 제외");
            int sent = 0;
            Action<object> send = item => { sent++; Check(File.ReadAllText(path).Contains("attempted"), "전송 전에 시도 상태 영속화"); };
            Func<object, string> ok = item => null, blocked = item => "창 없음";
            Call(q, "Process", "account", 20L, blocked, send);
            Check(sent == 0 && State(q, 0) == "blocked", "사전 검사 실패는 입력하지 않음");
            Call(q, "RetryBlocked"); Call(q, "Process", "account", 20L, ok, send);
            Check(sent == 1 && State(q, 0) == "submitted", "사용자 재시도 후 한 번 전송");
            q = New(path); Call(q, "Process", "account", 20L, ok, send); Check(sent == 1, "재시작 후 전송 중복 없음");
            Add(q, "two"); Action<object> fail = item => { sent++; throw new Exception("결과 모름"); };
            Call(q, "Process", "account", 20L, ok, fail); Check(State(q, 1) == "unknown", "불확실 결과 저장");
            Call(q, "RetryBlocked"); q = New(path); Call(q, "Process", "account", 20L, ok, send);
            Check(sent == 2 && State(q, 1) == "unknown", "불확실 결과는 재시도하지 않음");
            Add(q, "three"); Call(q, "Process", "account", 30L, ok, send);
            Check(State(q, 2) == "cancelled" && sent == 2, "대상 변경 시 이전 알림 전송 금지");
            Add(q, "four"); Items(q)[3].GetType().GetField("State").SetValue(Items(q)[3], "attempted"); Call(q, "Save");
            q = New(path); Check(State(q, 3) == "unknown", "전송 중 종료 시 재시작 재전송 금지");
            Call(q, "Enqueue", "long", "account", 20L, 10L, new string('가', 4000), 100L);
            Check(((string)Items(q)[4].GetType().GetField("Body").GetValue(Items(q)[4])).Length < 3500, "긴 이력 길이 제한");
            var db = (Database)Activator.CreateInstance(typeof(Database), true);
            var store = typeof(Database).GetField("Operations", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(db);
            var recordType = typeof(Bot).Assembly.GetType("KakaotalkBot.OperationRecord");
            var record = Activator.CreateInstance(recordType);
            foreach (var pair in new[] { new[] { "Kind", "memo" }, new[] { "Key", "memo:privacy" }, new[] { "Text", "비공개 본문" } }) recordType.GetField(pair[0]).SetValue(record, pair[1]);
            recordType.GetField("ChatId").SetValue(record, 10L); recordType.GetField("UserId").SetValue(record, 99L); recordType.GetField("At").SetValue(record, 200L);
            store.GetType().GetMethod("Put", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(store, new[] { record });
            var config = (Settings)Activator.CreateInstance(typeof(Settings), true); config.OperatorAlertsEnabled = true; config.OperatorAlertAccount = "account"; config.OperatorAlertRoomId = 20;
            var room = new ChatRoomInfo(); typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 10L, null); typeof(ChatRoomInfo).GetProperty("AccountPath").SetValue(room, "account", null);
            Call(q, "Collect", db, room, config);
            Check(!File.ReadAllText(path).Contains("비공개 본문"), "메모 본문 기본 비공개");
            Check(Items(q).Count == 6, "메모 알림 수집");
            Call(q, "Collect", db, room, config); Check(Items(q).Count == 6, "반복 수집 중복 없음");
            config.OperatorAlertRoomId = 10; Call(q, "Collect", db, room, config); Check(Items(q).Count == 6, "감시 방으로 알림 전송 금지");
            File.WriteAllText(Path.Combine(args[0], "bad.json"), "{}"); bool rejected = false;
            try { New(Path.Combine(args[0], "bad.json")); } catch { rejected = true; }
            Check(rejected, "손상된 전송 기록 차단");
            Console.WriteLine("PASS: " + checks + " notification checks (no real messages)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
