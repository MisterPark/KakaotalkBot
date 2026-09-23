using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using KakaotalkBot;

class RoomTitleTests
{
    static Assembly assembly = typeof(Bot).Assembly;
    static int checks;
    static object New(string name) { return Activator.CreateInstance(assembly.GetType("KakaotalkBot." + name), true); }
    static object Call(object target, string name, params object[] args)
    { return target.GetType().GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).Invoke(target, args); }
    static object Field(object target, string name) { return target.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target); }
    static int Active(object store, long chat, long user, long now) { return ((IEnumerable)Call(store, "Active", chat, user, now)).Cast<object>().Count(); }
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); checks++; }
    static void Fails(Action action)
    { try { action(); } catch (TargetInvocationException ex) { if (ex.InnerException is ArgumentException || ex.InnerException is FormatException || ex.InnerException is InvalidOperationException) { checks++; return; } throw; } throw new Exception("오류가 발생해야 합니다."); }
    static void AddMonthly(object ops, long id, long count, string kind = "monthly", long chat = 10)
    {
        var row = New("OperationRecord");
        foreach (var pair in new Dictionary<string, object> { { "Key", kind + ":" + chat + ":" + id }, { "Kind", kind }, { "Month", "2026-08" }, { "ChatId", chat }, { "UserId", id }, { "Count", count } })
            row.GetType().GetField(pair.Key).SetValue(row, pair.Value);
        Call(ops, "Put", row);
    }
    static int Main()
    {
        try
        {
            long now = new DateTimeOffset(2026, 9, 1, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            long next = new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            long id = 9007199254740993L;
            var store = New("RoomTitleStore");
            Check((bool)Call(store, "Grant", 10L, id, 7L, 1L, "토론왕", 0, now), "칭호 지정");
            Check(!(bool)Call(store, "Grant", 10L, id, 7L, 1L, "토론왕", 0, now), "동일 로그 중복 방지");
            Check(!(bool)Call(store, "Grant", 10L, id, 7L, 2L, "토론왕", 0, now), "활성 칭호 중복 방지");
            Check(Active(store, 20, id, now) == 0 && Active(store, 10, id, now) == 1, "방 격리");
            Call(store, "Grant", 10L, id, 7L, 3L, "기간제", 1, now);
            Check(Active(store, 10, id, now + 86399) == 2 && Active(store, 10, id, now + 86400) == 1, "만료 경계");
            Check((bool)Call(store, "Revoke", 10L, id, "토론왕", now + 1, 7L), "해제");
            Check(Active(store, 10, id, now + 1) == 1 && ((IDictionary)Field(store, "Rows")).Count == 2, "해제 이력 보존");
            var restored = New("RoomTitleStore");
            var rows = (List<List<object>>)Call(store, "ToRows");
            Call(restored, "Load", rows.Select(r => r.Select(Convert.ToString).ToList()).ToList());
            Check(Active(restored, 10, id, now + 1) == 1, "64비트 ID 저장 복원");
            Fails(() => Call(store, "Grant", 10L, id, 0L, 4L, "권한없음", 0, now));
            Fails(() => Call(store, "Grant", 10L, id, 7L, 4L, "줄\n바꿈", 0, now));
            var ops = New("OperationsStore");
            AddMonthly(ops, 1, 100); AddMonthly(ops, 2, 90); AddMonthly(ops, 3, 80); AddMonthly(ops, 4, 80); AddMonthly(ops, 5, 70);
            AddMonthly(ops, 6, 0, "popularity"); AddMonthly(ops, 7, -10, "popularity"); AddMonthly(ops, 8, 5, "popularity");
            var auto = New("RoomTitleStore");
            Call(auto, "AwardMonthly", ops, now);
            Check(((IDictionary)Field(auto, "Rows")).Count == 5, "월간 공동3위 및 양수 인기도");
            Call(auto, "AwardMonthly", ops, now + 1);
            Check(((IDictionary)Field(auto, "Rows")).Count == 5, "월간 중복 부여 방지");
            Check(Active(auto, 10, 1, next - 1) == 1 && Active(auto, 10, 1, next) == 0, "한국시간 월 만료");
            Call(auto, "Revoke", 10L, 1L, "월간 활동왕", now + 2, 7L);
            Call(auto, "AwardMonthly", ops, now + 3);
            Check(Active(auto, 10, 1, now + 3) == 0, "해제한 자동 칭호 재생성 금지");
            var policy = assembly.GetType("KakaotalkBot.OperatorCommandPolicy");
            Check((bool)policy.GetMethod("RequiresOperator").Invoke(null, new object[] { "/네임드지정 @유저 토론왕" }) &&
                (bool)policy.GetMethod("RequiresOperator").Invoke(null, new object[] { "/네임드해제 @유저 토론왕" }), "운영진 명령 등록");
            Fails(() => Call(Database.Instance, "ChangeNamedTitle", 10L, 7L, id, 99L, "권한없음", 0, false));
            var display = New("RoomTitleStore");
            typeof(Database).GetField("Titles", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(Database.Instance, display);
            var activity = Field(Database.Instance, "activity");
            long actual = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            for (int i = 1; i <= 21; i++)
            {
                Database.Instance.AddUser(9000 + i, "보존이름");
                Call(activity, "Observe", 10L, 9000L + i, "방닉" + i.ToString("D2"), true, actual, (long)i, "chat");
                Call(display, "Grant", 10L, 9000L + i, 7L, (long)i, "정보통", 0, actual);
            }
            string page1 = Database.Instance.NamedList(10), page2 = Database.Instance.NamedList(10, 2);
            Check(page1.Contains("방닉20") && !page1.Contains("방닉21") && page2.Contains("방닉21"), "20명 페이지 분리");
            Check(!page1.Contains("9001") && !page1.Contains("보존이름") && Database.Instance.NamedTitles(20, 9001) == "없음", "ID 비공개·방별 이름 격리");
            Call(activity, "Observe", 10L, 9001L, "방닉01", false, actual + 1, 100L, "event");
            Check(!Database.Instance.NamedList(10).Contains("방닉01") && Database.Instance.NamedTitles(10, 9001) == "정보통", "퇴장자 목록 제외·칭호 보존");
            Console.WriteLine("PASS: " + checks + " named title checks (no network writes or sends)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}