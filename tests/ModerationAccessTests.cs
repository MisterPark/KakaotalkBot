using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using KakaotalkBot;
using Newtonsoft.Json.Linq;

internal static class ModerationAccessTests
{
    private const long Super = 9007199254740993L;
    private const long Chat = 10, Session = 12;
    private const BindingFlags Flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
    private static readonly Assembly Assembly = typeof(Bot).Assembly;
    private static int passed;
    private static Type Type(string name) { return Assembly.GetType("KakaotalkBot." + name, true); }
    private static object New(string name) { return Activator.CreateInstance(Type(name), true); }
    private static void Set(object target, string name, object value) { target.GetType().GetField(name, Flags).SetValue(target, value); }
    private static object Get(object target, string name) { return target.GetType().GetField(name, Flags).GetValue(target); }
    private static object Call(object target, string name, params object[] args) { return target.GetType().GetMethod(name, Flags).Invoke(target, args); }
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); passed++; Console.WriteLine("PASS " + message); }
    private static bool Allowed(Bot bot, long actor, long chat = Chat, long session = Session, long target = 0, string nickname = null)
    { return (bool)Call(bot, "ModerationAllowed", chat, actor, session, target, nickname); }
    private static JObject State(string path) { return JObject.Parse(File.ReadAllText(path)); }
    private static int Queued(string path) { return ((JArray)State(path)["Rooms"][Chat.ToString()]["Moderation"]).Count; }
    private static object Member(long id, string name, int role)
    {
        var member = New("ChatMemberState");
        Set(member, "UserId", id); Set(member, "Nickname", name); Set(member, "MemberType", role); Set(member, "IsPresent", true);
        return member;
    }
    private static void SubmitWithoutInput(Bot bot, long actor, long log, long at)
    {
        var worker = Type("StaInputWorker").GetField("Instance", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
        var command = new Command { Keyword = "/내보낸목록", ChatId = Chat, AuthorId = actor, LogId = log, ReceivedAt = at, Nickname = "코드몽키" };
        Action build = () => Check((bool)Call(bot, "TryModerationCommand", command), "관리 명령 처리 경로 연결");
        // 실제 명령 생성과 기록까지 실행하되, UI 입력 단계는 항상 취소합니다.
        Check((bool)Call(worker, "TryBuildCommand", build, (Func<bool>)(() => false)), "입력 없는 명령 생성");
        worker.GetType().GetMethod("Invoke", Flags, null, new[] { typeof(Action) }, null).Invoke(worker, new object[] { (Action)(() => { }) });
    }
    private static void ElectionSettingWithoutInput(Bot bot, long actor, string argument, string verb = "/선거설정")
    {
        var worker = Type("StaInputWorker").GetField("Instance", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
        var command = new Command { Keyword = verb + " " + argument, ChatId = Chat, AuthorId = actor, ReceivedAt = DateTimeOffset.UtcNow.ToUnixTimeSeconds(), Nickname = "테스트" };
        Check((bool)Call(worker, "TryBuildCommand", (Action)(() => Call(bot, "TryGovernanceCommand", command)), (Func<bool>)(() => false)), "선거설정 처리 (채팅 전송 차단)");
        worker.GetType().GetMethod("Invoke", Flags, null, new[] { typeof(Action) }, null).Invoke(worker, new object[] { (Action)(() => { }) });
    }
    private static int Main()
    {
        try
        {
            // 운영 설정은 읽거나 복사하지 않습니다. 합성 키와 단 하나의 슈퍼계정을 사용합니다.
            var config = New("BotIdentityConfiguration");
            config.GetType().GetProperty("Key").SetValue(config, Convert.ToBase64String(Enumerable.Range(1, 64).Select(i => (byte)i).ToArray()), null);
            config.GetType().GetProperty("SpreadsheetId").SetValue(config, "local-test", null);
            var identity = Activator.CreateInstance(Type("BotIdentity"), Flags, null, new[] { config }, null);
            string encoded = (string)Call(identity, "Encode", Super);
            config.GetType().GetProperty("SuperUserId").SetValue(config, encoded, null);
            File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BotIdentity.json"), Newtonsoft.Json.JsonConvert.SerializeObject(config));
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cases", "governance.json");
            var store = Activator.CreateInstance(Type("GovernanceStore"), Flags, null, new object[] { path, null }, null);
            var db = Database.Instance;
            Set(db, "Governance", store); Set(db, "governanceIdentity", identity);
            var operators = Get(db, "Operators");
            var bot = (Bot)FormatterServices.GetUninitializedObject(typeof(Bot));
            Set(bot, "isBotRunning", true); Set(bot, "macroSession", Session);
            Set(bot, "moderationCompletions", new ConcurrentQueue<Action>());
            var room = new ChatRoomInfo();
            typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, Chat, null);
            typeof(ChatRoomInfo).GetProperty("LinkId").SetValue(room, 99L, null);
            typeof(Bot).GetProperty("SelectedRoom").SetValue(bot, room, null);
            long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();

            Call(bot, "RefreshModerationAccess");
            Check(Allowed(bot, Super), "슈퍼계정은 선거·명부·수동 봇 ID 설정 없이 목록 실행 허용");
            Check(!Allowed(bot, 7), "같은 닉네임의 다른 계정은 슈퍼권한 없음");
            Check(!Allowed(bot, 9), "이전 계정으로 슈퍼권한 확장 금지");
            SubmitWithoutInput(bot, Super, 101, now);
            Check(Queued(path) == 1 && !(bool)State(path)["Rooms"][Chat.ToString()]["Enabled"], "선거를 켜지 않고 관리 명령 기록");
            Check((string)State(path)["Rooms"][Chat.ToString()]["Moderation"][0]["Actor"] == encoded, "관리 이력에 고정 슈퍼 ID 저장");
            SubmitWithoutInput(bot, Super, 101, now);
            Check(Queued(path) == 1, "동일 메시지 중복 실행 방지");
            SubmitWithoutInput(bot, 7, 102, now);
            SubmitWithoutInput(bot, Super, 103, now - 60);
            Check(Queued(path) == 1, "사칭 계정과 오래된 명령은 기록·실행하지 않음");

            var roster = New("ChatRosterSnapshot");
            Set(roster, "ChatId", Chat); Set(roster, "LinkId", 99L); Set(roster, "IsOpenGroup", true); Set(roster, "ObservedAt", now);
            var members = (IList)Get(roster, "Members");
            members.Add(Member(Super, "코드몽키", 2)); members.Add(Member(7, "코드몽키", 2));
            members.Add(Member(20, "봇", 4)); members.Add(Member(30, "방장", 1)); members.Add(Member(40, "부방장", 4));
            members.Add(Member(50, "대상", 2)); members.Add(Member(51, "중복", 2)); members.Add(Member(52, "중복", 2));
            Call(operators, "Observe", roster); Call(bot, "ObserveGovernance", roster); Call(bot, "RefreshModerationAccess");
            Check(ReferenceEquals(Get(bot, "moderationRoster"), roster), "봇 ID 미입력 상태에서도 관리 명부 전달");
            Check(Allowed(bot, Super), "일반 참여자인 슈퍼계정의 관리 권한 유지");
            Check(!Allowed(bot, 30), "일반 운영진은 봇의 실제 관리 권한 확인 필요");
            ElectionSettingWithoutInput(bot, Super, "켜기");
            Check(!(bool)Call(store, "Enabled", Chat), "본인 ID 미확인 상태에서는 잘못된 선거 명부를 만들지 않음");
            Set(roster, "OwnAuthorId", 20L); Call(bot, "ObserveGovernance", roster); Call(bot, "RefreshModerationAccess");
            Check(bot.OwnAuthorId == 0 && Allowed(bot, 30), "수동 ID 입력 없이 자동 확인한 봇 계정으로 관리 권한 확인");
            var liveRoster = ((IDictionary)Get(store, "live"))[Chat];
            var electors = ((IEnumerable)Get(liveRoster, "Present")).Cast<string>().ToArray();
            Check(!electors.Contains((string)Call(db, "GovernanceUser", 20L)) && electors.Contains(encoded), "자동 확인한 봇은 선거 명부에서 제외하고 슈퍼관리자는 유지");
            ElectionSettingWithoutInput(bot, 7, "켜기");
            Check(!(bool)Call(store, "Enabled", Chat), "같은 닉네임의 일반 계정은 선거 설정을 변경하지 못함");
            ElectionSettingWithoutInput(bot, Super, "켜기");
            Check((bool)Call(store, "Enabled", Chat), "슈퍼관리자는 수동 봇 ID 없이 선거를 켤 수 있음");
            ElectionSettingWithoutInput(bot, 7, "", "/선거시작"); ElectionSettingWithoutInput(bot, 7, "", "/보궐선거");
            Check(((JArray)State(path)["Rooms"][Chat.ToString()]["Elections"]).Count == 0, "일반 사용자는 정기·보궐선거를 강제로 열 수 없음");
            Set(roster, "OwnAuthorId", 0L);
            ElectionSettingWithoutInput(bot, Super, "끄기");
            Check(!(bool)Call(store, "Enabled", Chat), "봇 ID 확인 대기 중에도 권한 있는 사용자는 선거를 끌 수 있음");
            Set(roster, "OwnAuthorId", 20L); Set(roster, "ChatId", Chat + 1); Call(bot, "RefreshModerationAccess");
            Check((long)bot.GetType().GetProperty("CurrentOwnAuthorId", Flags).GetValue(bot, null) == 0, "다른 방의 자동 확인 ID를 사용하지 않음");
            Set(roster, "ChatId", Chat); Set(roster, "ObservedAt", now - 60);
            Check((long)bot.GetType().GetProperty("CurrentOwnAuthorId", Flags).GetValue(bot, null) == 0, "만료된 자동 확인 ID를 사용하지 않음");
            Set(roster, "ObservedAt", now);
            bot.OwnAuthorId = 20; Call(bot, "RefreshModerationAccess");
            Check(Allowed(bot, 30) && Allowed(bot, 40), "선거가 꺼져 있어도 방장·부방장 허용");
            Check(!Allowed(bot, 7), "최신 명부에서도 닉네임 사칭 거부");
            Check(!Allowed(bot, Super, Chat, Session, Super, "코드몽키") && !Allowed(bot, 30, Chat, Session, Super, "코드몽키"), "슈퍼계정 내보내기 보호");
            Check(!Allowed(bot, Super, Chat, Session, 30, "방장") && !Allowed(bot, Super, Chat, Session, 20, "봇"), "방장·봇 내보내기 보호");
            Check(!Allowed(bot, Super, Chat, Session, 51, "중복") && !Allowed(bot, Super, Chat, Session, 999, "없음"), "중복 닉네임·미참여 대상 거부");
            Check(Allowed(bot, Super, Chat, Session, 50, "대상"), "검증된 일반 참여자 대상 허용");
            Call(operators, "Invalidate", Chat, 30L, 2L, 2); Call(bot, "RefreshModerationAccess");
            Check(!Allowed(bot, Super, Chat, Session, 50, "대상"), "역할 변경 확인 중에는 내보낼 대상 보호");
            Check(Allowed(bot, Super), "역할 변경 중에도 슈퍼 목록 조회 허용");
            Set(roster, "ObservedAt", now - 60); Call(bot, "RefreshModerationAccess");
            Check(Allowed(bot, Super) && !Allowed(bot, 30) && !Allowed(bot, Super, Chat, Session, 50, "대상"), "오래된 명부는 슈퍼 조회만 허용");
            Check(!Allowed(bot, Super, Chat + 1) && !Allowed(bot, Super, Chat, Session - 1), "다른 방·이전 수신 세션 거부");
            Set(bot, "isBotRunning", false);
            Check(!Allowed(bot, Super), "봇 중지 시 슈퍼 매크로도 중단");
            Console.WriteLine("PASS: " + passed + " moderation checks (no chat, Sheets, or UI input)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error.GetBaseException()); return 1; }
    }
}
