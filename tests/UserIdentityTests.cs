using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using KakaotalkBot;

internal static class UserIdentityTests
{
    private const long A = 8172472438151128507L, B = 8172472438151128508L;
    private static int passed;
    private static readonly string[] Headers = { "nickname", "take_attendance", "attendance_at", "point", "popularity", "password", "contribution", "user_id", "leave_count", "kick_count" };

    private static int Main()
    {
        try
        {
            Run("64비트 ID 문자열 저장과 재조회", () =>
            {
                foreach (long id in new[] { A, B, long.MaxValue })
                {
                    var original = new User(id) { Nickname = "동일닉", Password = "0012", Point = 42, LeaveCount = 2, KickCount = 3 };
                    var row = original.ToRow();
                    Check(row[7] is string && (string)row[7] == id.ToString(CultureInfo.InvariantCulture), "ID 숫자 변환");
                    var restored = User.ToUser(row.Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).ToList());
                    Check(restored.UserId == id && restored.Password == "0012" && restored.Point == 42 && restored.LeaveCount == 2 && restored.KickCount == 3, "왕복 손실");
                }
            });
            Run("닉네임 변경·도용 후에도 계정 분리", () =>
            {
                var db = FreshDatabase();
                var original = db.AddUser(A, "기존닉"); original.Point = 100; original.Password = "owner"; original.Contribution = 7;
                Check(ReferenceEquals(original, db.AddUser(A, "변경닉")), "동일 ID 중복 등록");
                var other = db.AddUser(B, "변경닉");
                Check(original.Point == 100 && original.Password == "owner" && original.Contribution == 7, "변경 후 정보 손실");
                Check(other.Point == 0 && other.Contribution == 0 && db.UserTable.Count == 2, "닉네임으로 자산 상속");
                User ignored;
                Check(!db.FindUser(0, out ignored), "0 ID 조회");
                Throws(() => db.AddUser(0, "이름"));
            });
            Run("출석은 ID별로 하루 한 번", () =>
            {
                var db = FreshDatabase();
                Check(!db.CheckAttendance(A, "이름"), "첫 출석");
                Check(db.CheckAttendance(A, "다른이름"), "닉변 중복 출석");
                Check(!db.CheckAttendance(B, "다른이름"), "동명이인 출석");
                Check(db.UserTable.All(user => user.Point == 10), "출석 포인트 분리");
                db.ResetAttendance();
                Check(db.CheckAttendance(A, "재시작후") && db.UserTable[0].Point == 10, "재시작 중복 출석");
                db.UserTable[0].AttendanceAt = DateTime.Today.AddDays(-1);
                Check(!db.CheckAttendance(A, "다음날") && db.UserTable[0].Point == 20, "다음날 출석");
            });
            Run("미등록 멘션 대상 조회와 기존 정보 보존", () =>
            {
                var db = FreshDatabase(); long id; int amount; string nickname;
                var command = new Command { AuthorId = A, Keyword = "/조회 @새 사용자", Mentions = new[] { new ChatMention(B, new[] { 1 }, 5) } };
                Check(command.TryReadTarget("/조회", false, out id, out amount, out nickname), "조회 대상 파싱");
                var user = db.GetOrAddUser(id, nickname);
                Check(user.UserId == B && user.Nickname == "새 사용자" && user.Point == 0 && user.Popularity == 0 && user.Contribution == 0, "자동 등록 초기값");
                user.Point = 25; user.Password = "saved";
                Check(ReferenceEquals(user, db.GetOrAddUser(B, "다른표시")) && db.UserTable.Count == 1, "조회 중 중복 등록");
                Check(user.Nickname == "새 사용자" && user.Point == 25 && user.Password == "saved", "멘션으로 기존 정보 덮어쓰기");
                command.Mentions = new[] { new ChatMention(B, new[] { 1 }, 5, "방 프로필") };
                Check(command.TryReadTarget("/조회", false, out id, out amount, out nickname) && nickname == "방 프로필", "프로필 이름 우선 적용");
                Throws(() => db.GetOrAddUser(0, "무효"));
                Check(db.UserTable.Count == 1, "무효 ID 등록");
            });
            Run("미등록 대상에게 좋아·싫어 처리", () =>
            {
                foreach (bool increase in new[] { true, false })
                {
                    var db = FreshDatabase(); var author = db.AddUser(A, "작성자"); author.Point = 20;
                    string error; User target;
                    Check(db.ChangePopularity(A, B, 10, increase, out error), "미등록 대상 거래 거부");
                    Check(db.FindUser(B, out target) && target.Popularity == (increase ? 10 : -10) && target.Point == 0 && author.Point == 10, "자동 등록 후 거래 반영");
                    Check(db.UserTable.Count == 2, "대상 등록 횟수");
                }
            });
            Run("미등록 작성자는 등록 후 잔액 검사", () =>
            {
                var db = FreshDatabase(); string error;
                Check(!db.ChangePopularity(A, B, 10, true, out error) && error.StartsWith("포인트가 부족"), "새 작성자의 잔액 검사");
                Check(db.UserTable.Count == 2 && db.UserTable.All(user => user.Point == 0 && user.Popularity == 0), "실패한 거래의 자산 변경");
                db = FreshDatabase();
                Check(!db.ChangePopularity(0, B, 10, true, out error) && db.UserTable.Count == 0, "잘못된 ID 자동 등록");
            });
            Run("동명이인 인기도 거래와 자기 자신 판정", () =>
            {
                var db = FreshDatabase(); var author = db.AddUser(A, "같은이름"); var target = db.AddUser(B, "같은이름"); author.Point = 20;
                string error;
                Check(db.ChangePopularity(A, B, 10, true, out error), "동명이인 거래 거부");
                Check(author.Point == 10 && target.Popularity == 10 && target.Point == 0, "잘못된 자산 변경");
                db.AddUser(A, "다른이름");
                Check(!db.ChangePopularity(A, A, 1, true, out error), "닉변 자기 거래");
                Check(db.ChangePopularity(A, B, 5, false, out error) && target.Popularity == 5 && author.Point == 5, "싫어요 ID 처리");
            });
            Run("잘못된 금액·잔액 부족·오버플로는 차감하지 않음", () =>
            {
                var db = FreshDatabase(); var author = db.AddUser(A, "a"); var target = db.AddUser(B, "b"); author.Point = 20;
                string error;
                foreach (int amount in new[] { 0, -1, int.MinValue, 21 })
                    Check(!db.ChangePopularity(A, B, amount, true, out error) && author.Point == 20, "잘못된 차감");
                target.Popularity = int.MaxValue;
                Check(!db.ChangePopularity(A, B, 1, true, out error) && author.Point == 20, "오버플로 부분 반영");
            });
            Run("멘션 ID로 대상을 읽고 닉네임 문자열은 거부", () =>
            {
                long id; int amount;
                var command = new Command { AuthorId = A, Keyword = "/좋아 @동일닉 15", Mentions = new[] { new ChatMention(B, new[] { 1 }, 3) } };
                Check(command.TryReadTarget("/좋아", true, out id, out amount) && id == B && amount == 15, "대상/금액");
                command.Mentions = null;
                Check(!command.TryReadTarget("/좋아", true, out id, out amount), "가짜 @ 허용");
                command.Mentions = new[] { new ChatMention(B, new[] { 1 }, 3) };
                foreach (string suffix in new[] { " -1", " 0", " 2147483648", " 텍스트" })
                {
                    command.Keyword = "/좋아 @동일닉" + suffix;
                    Check(!command.TryReadTarget("/좋아", true, out id, out amount), "금액 검증");
                }
                command.Keyword = "/조회 @동일닉";
                Check(command.TryReadTarget("/조회", false, out id, out amount) && id == B, "멘션 조회");
                command.Mentions = new[] { new ChatMention(B, new[] { 1, 2 }, 3) };
                Check(!command.TryReadTarget("/조회", false, out id, out amount), "중복 멘션");
                command.Mentions = new[] { new ChatMention(B, new[] { 1 }, int.MaxValue) };
                Check(!command.TryReadTarget("/조회", false, out id, out amount), "길이 오버플로");
            });
            Run("공백과 숫자가 포함된 멘션 이름", () =>
            {
                long id; int amount;
                var command = new Command { AuthorId = A, Keyword = "/좋아 @사용자 123 4", Mentions = new[] { new ChatMention(B, new[] { 1 }, 7) } };
                Check(command.TryReadTarget("/좋아", true, out id, out amount) && id == B && amount == 4, "이름 숫자를 금액으로 오인");
                command.Keyword = "/좋아 @사용자 123";
                Check(command.TryReadTarget("/좋아", true, out id, out amount) && amount == 10, "기본 금액");
            });
            Run("시트 헤더·문자열 ID·중복 검사", () =>
            {
                var rows = new List<IList<object>> { Headers.Cast<object>().ToList(), new User(A).ToRow(), new User(B).ToRow() };
                Check(Parse(rows).Count == 2, "정상 시트");
                rows[2][7] = rows[1][7]; Throws(() => Parse(rows));
                rows.RemoveAt(2); rows[1][7] = A; Throws(() => Parse(rows));
                foreach (string invalid in new[] { "", "0", "-1", "8.17247243815113E+18", "9223372036854775808" })
                { rows[1][7] = invalid; Throws(() => Parse(rows)); }
                Throws(() => User.ToUser(new User(A).ToRow().Take(7).Select(value => Convert.ToString(value)).ToList()));
                Check(Parse(new List<IList<object>> { Headers.Cast<object>().ToList() }).Count == 0, "새 DB 초기 상태");
                Throws(() => Parse(new List<IList<object>>()));
            });
            Run("기존 ID 계정의 빈 횟수 컬럼과 잘못된 값 검사", () =>
            {
                var row = new User(A) { Point = 42 }.ToRow().Take(8).Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).ToList();
                var user = User.ToUser(row);
                Check(user.Point == 42 && user.LeaveCount == 0 && user.KickCount == 0, "기존 계정 보존");
                Check(Parse(new List<IList<object>> { Headers.Cast<object>().ToList(), row.Cast<object>().ToList() }).Count == 1, "빈 새 컬럼 읽기");
                row.Add(""); row.Add("");
                Check(User.ToUser(row).KickCount == 0, "빈 문자열 횟수 읽기");
                foreach (string invalid in new[] { "-1", "1.5", "2147483648" })
                { row[8] = invalid; Throws(() => User.ToUser(row)); }
            });
            Run("퇴장·강퇴를 ID별로 분리하고 이벤트 중복 차단", () =>
            {
                var db = FreshDatabase();
                Check(db.RecordDeparture(A, "동일닉", 10, 100, false), "첫 퇴장 기록");
                Check(db.RecordDeparture(B, "동일닉", 10, 100, false), "여러 퇴장 대상 분리");
                Check(!db.RecordDeparture(A, "동일닉", 10, 100, false), "동일 이벤트 중복");
                Check(db.RecordDeparture(A, "변경닉", 10, 101, true), "강퇴 기록");
                Check(db.RecordDeparture(A, "변경닉", 10, 102, false), "재입장 후 퇴장 기록");
                var user = db.GetOrAddUser(A);
                Check(user.LeaveCount == 2 && user.KickCount == 1 && db.ChatUserName(10, A) == "변경닉", "ID별 횟수");
                Check(db.GetOrAddUser(B).LeaveCount == 1 && db.GetOrAddUser(B).KickCount == 0, "동명이인 횟수 혼합");
                Throws(() => db.RecordDeparture(0, "무효", 10, 103, false));
            });
            Run("수신 이벤트 집계와 일반 문구·다른 방 제외", TestDepartureEvents);
            Run("미등록 사용자의 첫 강퇴를 해석·등록·저장하고 재입장 후 유지", TestFirstKick);
            Run("채팅·입퇴장 수신에서 경험치·상태·닉네임 이력을 함께 갱신", TestActivityIntegration);
            Run("로드하지 못한 DB는 저장 차단", () => { Throws(() => FreshDatabase().UpdateUserTable()); });
            Run("수신·퀴즈 큐·기여도에 작성자 ID 유지", TestIncoming);
            Run("암호 변경에서 동명이인·닉변·만료 처리", TestPasswords);
            Run("1:1 방은 요청한 ID와 유일하게 일치할 때만 선택", TestDirectRoom);
            Console.WriteLine("PASS: " + passed + " identity cases (no network writes or chat sends)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void TestIncoming()
    {
        var db = FreshDatabase(); var bot = BareBot();
        Set(bot, "chatLog", new List<string>()); Set(bot, "commands", new Queue<Command>()); Set(bot, "quizAnswers", new Queue<Bot.QuizAnswer>());
        var room = new ChatRoomInfo(); typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 10L, null);
        typeof(Bot).GetProperty("SelectedRoom").SetValue(bot, room, null);
        Type type = typeof(Bot).Assembly.GetType("KakaotalkBot.ChatMessage");
        Action<long, string, long> deliver = (id, name, chat) =>
        {
            object message = Activator.CreateInstance(type, true);
            type.GetField("AuthorId").SetValue(message, id); type.GetField("Nickname").SetValue(message, name);
            type.GetField("ChatId").SetValue(message, chat); type.GetField("Message").SetValue(message, "정답");
            Invoke(bot, "HandleIncomingMessage", message);
        };
        deliver(A, "같은이름", 10); deliver(B, "같은이름", 10); deliver(A, "변경닉", 10); deliver(0, "변경닉", 10); deliver(99, "변경닉", 999);
        User a, b; db.FindUser(A, out a); db.FindUser(B, out b);
        Check(db.UserTable.Count == 2 && a.Contribution == 2 && b.Contribution == 1 && db.ChatUserName(10, A) == "변경닉", "수신 ID 격리");
        var queue = (Queue<Bot.QuizAnswer>)Get(bot, "quizAnswers");
        Check(queue.Select(answer => answer.AuthorId).SequenceEqual(new[] { A, B, A }), "퀴즈 ID 전달");
    }

    private static void TestDepartureEvents()
    {
        var db = FreshDatabase(); var bot = BareBot();
        var commands = new Queue<Command>(); var answers = new Queue<Bot.QuizAnswer>();
        Set(bot, "chatLog", new List<string>()); Set(bot, "commands", commands); Set(bot, "quizAnswers", answers);
        var room = new ChatRoomInfo(); typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 10L, null);
        typeof(Bot).GetProperty("SelectedRoom").SetValue(bot, room, null);
        Type type = typeof(Bot).Assembly.GetType("KakaotalkBot.ChatMessage");
        var eventField = type.GetField("RoomEvent");
        Action<long, long, string> deliver = (chatId, logId, kind) =>
        {
            object message = Activator.CreateInstance(type, true);
            type.GetField("AuthorId").SetValue(message, A); type.GetField("ChatId").SetValue(message, chatId);
            type.GetField("LogId").SetValue(message, logId);
            type.GetField("Message").SetValue(message, "사용자님을 내보냈습니다.");
            type.GetField("EventCommand").SetValue(message, kind == "None" ? null : kind == "Join" ? "/입장" : "/퇴장");
            eventField.SetValue(message, Enum.Parse(eventField.FieldType, kind));
            Invoke(bot, "HandleIncomingMessage", message);
        };
        deliver(10, 100, "Leave"); deliver(10, 100, "Leave"); deliver(10, 101, "Kick");
        deliver(10, 102, "None"); deliver(999, 103, "Kick"); deliver(10, 104, "Join");
        var user = db.GetOrAddUser(A);
        Check(user.LeaveCount == 1 && user.KickCount == 1, "수신 이벤트 누락 또는 과다 집계");
        Check(commands.Count == 3 && answers.Count == 1 && user.Contribution == 1, "시스템 이벤트 기여도·퀴즈 제외");
        Check(!string.IsNullOrWhiteSpace(user.Nickname), "퇴장 사용자 이름 없음 처리");
        Check(bot.HasPendingUserSave, "횟수 변경 직후 저장 예약 누락");
        int saves = 0;
        Action save = () =>
        {
            saves++;
            var restored = User.ToUser(user.ToRow().Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).ToList());
            Check(restored.UserId == A && restored.LeaveCount == 1 && restored.KickCount == 1, "미등록 사용자와 두 횟수가 함께 저장되지 않음");
        };
        Invoke(bot, "FlushPendingUserChanges", save);
        Check(saves == 1 && !bot.HasPendingUserSave, "수신 묶음 저장 완료 처리");
        deliver(10, 100, "Leave");
        Invoke(bot, "FlushPendingUserChanges", save);
        Check(saves == 1, "중복 이벤트가 다시 저장됨");
        deliver(10, 105, "Leave");
        Throws(() => Invoke(bot, "FlushPendingUserChanges", (Action)(() => { throw new InvalidOperationException("저장 실패 모의"); })));
        Check(bot.HasPendingUserSave && user.LeaveCount == 2, "실패 시 저장 대기·횟수 소실");
        Invoke(bot, "FlushPendingUserChanges", (Action)(() => { saves++; }));
        Check(!bot.HasPendingUserSave && saves == 2 && user.LeaveCount == 2, "재시도 후 중복 집계");
    }

    private static void TestFirstKick()
    {
        var db = FreshDatabase(); var bot = BareBot();
        Set(bot, "chatLog", new List<string>()); Set(bot, "commands", new Queue<Command>()); Set(bot, "quizAnswers", new Queue<Bot.QuizAnswer>());
        var room = new ChatRoomInfo(); typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 10L, null);
        typeof(Bot).GetProperty("SelectedRoom").SetValue(bot, room, null);
        var assembly = typeof(Bot).Assembly;
        Type directoryType = assembly.GetType("KakaotalkBot.ChatUserDirectory");
        object directory = Activator.CreateInstance(directoryType, true);
        ((HashSet<long>)directoryType.GetField("OwnIds").GetValue(directory)).Add(9L);
        MethodInfo decode = assembly.GetType("KakaotalkBot.ChatMessageDecoder").GetMethod("Decode");
        object[] row = { 123L, 9L, 0L, 100L, "{\"feedType\":6,\"member\":{\"userId\":7160322422266675266,\"nickName\":\"케이\"}}", 0L, 0L };
        Action receive = () =>
        {
            var messages = ((IEnumerable)decode.Invoke(null, new object[] { room, row, directory, false })).Cast<object>().ToArray();
            Check(messages.Length == 1, "프로필과 시트에 없는 강퇴 대상 해석 실패");
            Invoke(bot, "HandleIncomingMessage", messages[0]);
        };
        receive(); receive();
        User user;
        Check(db.FindUser(7160322422266675266L, out user) && user.KickCount == 1 && user.LeaveCount == 0, "첫 강퇴 등록 또는 중복 제외 실패");
        Invoke(bot, "FlushPendingUserChanges", (Action)(() =>
        {
            var restored = User.ToUser(user.ToRow().Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).ToList());
            Check(restored.KickCount == 1 && restored.Nickname == "케이", "첫 강퇴 저장 실패");
        }));
        row[0] = 124L;
        row[4] = "{\"feedType\":4,\"members\":[{\"userId\":7160322422266675266,\"nickName\":\"변경닉\"}]}";
        receive();
        Check(db.UserTable.Count == 1 && user.KickCount == 1 && db.ChatUserName(10, user.UserId) == "변경닉", "재입장 시 기존 횟수 초기화");
    }

    private static void TestActivityIntegration()
    {
        var db = FreshDatabase(); var bot = BareBot();
        Set(bot, "chatLog", new List<string>()); Set(bot, "commands", new Queue<Command>()); Set(bot, "quizAnswers", new Queue<Bot.QuizAnswer>());
        var room = new ChatRoomInfo(); typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 10L, null);
        typeof(Bot).GetProperty("SelectedRoom").SetValue(bot, room, null);
        Type type = typeof(Bot).Assembly.GetType("KakaotalkBot.ChatMessage");
        var eventField = type.GetField("RoomEvent");
        Action<long, long, string, string, string> deliver = (log, at, text, kind, name) =>
        {
            object message = Activator.CreateInstance(type, true);
            type.GetField("AuthorId").SetValue(message, A); type.GetField("ChatId").SetValue(message, 10L);
            type.GetField("LogId").SetValue(message, log); type.GetField("SendAt").SetValue(message, at);
            type.GetField("Message").SetValue(message, text); type.GetField("Nickname").SetValue(message, name);
            if (kind != "None") type.GetField("EventCommand").SetValue(message, kind == "Join" ? "/입장" : "/퇴장");
            eventField.SetValue(message, Enum.Parse(eventField.FieldType, kind));
            Invoke(bot, "HandleIncomingMessage", message);
        };
        deliver(10, 100, "대화", "None", "처음");
        deliver(11, 160, "/조회", "None", "처음");
        Check(db.GetOrAddUser(A).Experience == 1, "명령어 경험치 제외");
        deliver(12, 160, "대화2", "None", "처음");
        deliver(13, 170, "/퇴장", "Kick", "처음");
        Check(db.FindRoomUser(10, A).IsPresent == false && db.GetOrAddUser(A).KickCount == 1, "강퇴 상태와 횟수");
        deliver(14, 180, "/입장", "Join", "변경");
        deliver(15, 219, "짧은 간격", "None", "변경");
        deliver(16, 220, "60초 경과", "None", "변경");
        Check(db.GetOrAddUser(A).Experience == 3 && db.GetOrAddUser(A).Level == 1, "경험치 지급 간격 또는 시스템 이벤트 제외");
        Check(db.FindRoomUser(10, A).IsPresent == true && db.RoomEvents.Count == 2 && db.NicknameHistory.Count == 1, "재입장 상태와 닉네임 이력");
        Check(db.HasPendingActivity && bot.HasPendingUserSave, "정상 종료 시 활동 저장 필요 표시");
        Invoke(bot, "FlushPendingUserChanges", (Action)(() => { }));
        Check(!bot.HasPendingUserSave && !db.HasPendingActivity, "활동 저장 완료 표시");
        Check(db.DescribeActivity(10, A).Contains("처음 → 변경") && db.DescribeActivity(10, A).Contains("참여 중"), "조회 이력 표시");
        Check(db.FindRoomUser(999, A) == null, "확인하지 않은 방 상태는 미확인");
    }

    private static void TestPasswords()
    {
        var db = FreshDatabase(); var owner = db.AddUser(A, "같은이름"); var other = db.AddUser(B, "같은이름"); owner.Password = "old";
        var bot = BareBot();
        Type requestType = typeof(Bot).GetNestedType("PasswordRequest", BindingFlags.NonPublic);
        Type stageType = typeof(Bot).GetNestedType("PasswordStage", BindingFlags.NonPublic);
        var dictionary = (IDictionary)Activator.CreateInstance(typeof(Dictionary<,>).MakeGenericType(typeof(long), requestType));
        Set(bot, "passwordRequests", dictionary);
        object request = Activator.CreateInstance(requestType, true);
        requestType.GetField("Stage", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(request, Enum.Parse(stageType, "CheckOld"));
        requestType.GetField("ExpiresAt", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(request, DateTime.UtcNow.AddMinutes(5));
        dictionary.Add(A, request);
        Func<long, string, object> receive = (id, text) => Invoke(bot, "ProcessDirectMessage", new Bot.Chat { AuthorId = id, Nickname = "같은이름", Message = text });
        Check(receive(B, "old") == null && dictionary.Contains(A), "동명이인이 인증 상태 소비");
        db.AddUser(A, "변경닉");
        Check(receive(A, "old") != null, "닉변 본인 인증 실패");
        Check(receive(B, "stolen") == null && owner.Password == "old", "다른 ID로 암호 변경");
        Check(receive(A, "new") != null && owner.Password == "new" && other.Password == "0000", "본인 ID 암호 변경");
        dictionary.Add(A, request);
        requestType.GetField("ExpiresAt", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(request, DateTime.UtcNow.AddSeconds(-1));
        Check(receive(A, "expired") == null && owner.Password == "new", "만료 후 암호 변경");
    }

    private static void TestDirectRoom()
    {
        var room = new ChatRoomInfo();
        typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 20L, null);
        typeof(ChatRoomInfo).GetProperty("DirectUserId").SetValue(room, B, null);
        typeof(ChatRoomInfo).GetProperty("Kind").SetValue(room, "DirectChat", null);
        typeof(ChatRoomInfo).GetProperty("Name").SetValue(room, "같은이름", null);
        Type receiver = typeof(Bot).Assembly.GetType("KakaotalkBot.ChatDatabaseReceiver");
        var find = receiver.GetMethod("FindDirectRoom", BindingFlags.NonPublic | BindingFlags.Static);
        Check(find.Invoke(null, new object[] { new[] { room }, A, 10L }) == null, "닉네임 1:1 방 허용");
        Check(ReferenceEquals(find.Invoke(null, new object[] { new[] { room }, B, 10L }), room), "ID 1:1 방 선택");
        Check(find.Invoke(null, new object[] { new[] { room, room }, B, 10L }) == null, "중복 1:1 방 선택");
    }
    private static IList Parse(IList<IList<object>> rows) { return (IList)typeof(GoogleSheetHelper).GetMethod("ParseUserRows", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new object[] { rows }); }
    private static Database FreshDatabase() { typeof(Database).GetField("instance", BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, null); return Database.Instance; }
    private static Bot BareBot() { return (Bot)FormatterServices.GetUninitializedObject(typeof(Bot)); }
    private static object Invoke(object target, string method, params object[] args) { return target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args); }
    private static void Set(object target, string field, object value) { target.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value); }
    private static object Get(object target, string field) { return target.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
    private static void Run(string name, Action test) { test(); passed++; Console.WriteLine("PASS: " + name); }
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    private static void Throws(Action action)
    {
        try { action(); }
        catch (Exception error)
        {
            Exception cause = error is TargetInvocationException ? error.InnerException : error;
            if (cause is ArgumentException || cause is FormatException || cause is InvalidOperationException) return;
            throw;
        }
        throw new Exception("예상한 오류가 발생하지 않았습니다.");
    }
}
