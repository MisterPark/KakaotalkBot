using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using KakaotalkBot;

class RoomOperatorTests
{
    static int count;
    static void Check(bool value, string label) { if (!value) throw new Exception(label); count++; }
    static ChatRosterSnapshot Roster(long time, long log, int role = 4)
    {
        var r = new ChatRosterSnapshot { ChatId = 10, LinkId = 99, IsOpenGroup = true, ObservedAt = time, LogId = log };
        r.Members.Add(new ChatMemberState { UserId = 9007199254740993L, Nickname = "동일 이름", IsPresent = true, MemberType = role, MemberPrivilege = "18446744073709551615" });
        r.Members.Add(new ChatMemberState { UserId = 20, Nickname = "동일 이름", IsPresent = true, MemberType = 2 });
        return r;
    }
    static int Main()
    {
        try
        {
            long id = 9007199254740993L;
            string error;
            var store = new RoomOperatorStore();
            Check(!store.Check(10, id, 100, out error), "수신 전 권한 없음");
            store.Observe(Roster(100, 1));
            Check(!store.CheckQuizIssuer(10, id, 0, 100, 100, out error), "부방장 퀴즈 등록 거부");
            Check(!store.CheckQuizIssuer(10, 20, id, 100, 100, out error), "동일 닉네임 슈퍼계정 사칭 거부");
            Check(store.CheckQuizIssuer(10, 20, 20, 100, 100, out error), "고정 ID 슈퍼계정 허용");
            Check(!store.CheckQuizIssuer(10, 20, 20, 116, 100, out error), "지연된 슈퍼계정 명령 거부");
            Check(store.Check(10, id, 100, out error), "부방장 허용");
            Check(!store.Check(10, 20, 100, out error), "동명 일반 사용자 거부");
            Check(!store.Check(11, id, 100, out error), "다른 방 권한 거부");
            Check(!store.Check(10, id, 116, out error), "정보 만료 거부");
            Check(!store.Check(10, id, 99, out error), "미래 시각 거부");
            var saved = store.Records.Values.Select(r => r.ToRow().Select(v => Convert.ToString(v, CultureInfo.InvariantCulture)).ToList()).ToList();
            var restored = new RoomOperatorStore(); restored.Load(saved);
            Check(restored.Records.Values.Single().UserId == id && restored.Records.Values.Single().Privilege == "18446744073709551615", "긴 ID와 권한 값 보존");
            Check(!restored.Check(10, id, 100, out error), "저장된 운영진만으로 권한 부여 금지");
            store.Invalidate(10, id, 2, 2);
            Check(!store.Check(10, id, 100, out error), "해임 수신 즉시 차단");
            store.Observe(Roster(101, 3));
            Check(!store.Check(10, id, 101, out error), "프로필 해임 반영 지연 시 차단 유지");
            store.Observe(Roster(102, 4, 2));
            Check(!store.Check(10, id, 102, out error) && !store.Records.Values.Single().IsOperator, "해임 저장 반영");
            store.Observe(Roster(103, 5, 1));
            Check(store.Check(10, id, 103, out error), "방장 허용");
            Check(store.CheckQuizIssuer(10, id, 0, 103, 103, out error), "방장 퀴즈 등록 허용");
            Check(!store.CheckQuizIssuer(10, id, 0, 120, 120, out error), "오래된 방장 정보 거부");
            store.Observe(Roster(104, 6, 16));
            Check(!store.Check(10, id, 104, out error), "미확인 역할 권한 없음");
            var absent = Roster(105, 7); absent.Members[0].IsPresent = false; store.Observe(absent);
            Check(!store.Check(10, id, 105, out error), "퇴장 운영진 권한 없음");
            store.Observe(Roster(106, 8)); store.Invalidate(10, 0, 9, 2);
            Check(!store.Check(10, id, 106, out error), "대상 미상 권한 변경 차단");
            store.ResetLive();
            Check(!store.Check(10, id, 106, out error), "수신 중지 시 권한 제거");
            var command = new Command { AuthorId = id, ChatId = 10, Keyword = "/메모 @이름 공백 사유 있는 내용",
                Mentions = new[] { new ChatMention(20, new[] { 1 }, 5) } };
            long target; string reason;
            Check(command.TryReadMentionText("/메모", out target, out reason) && target == 20 && reason == "사유 있는 내용", "공백 닉네임과 사유 파싱");
            command.Keyword = "/메모 @이름 공백";
            Check(command.TryReadMentionText("/메모", out target, out reason) && reason == "", "사유 생략");
            command.Mentions = null;
            Check(!command.TryReadMentionText("/메모", out target, out reason), "가짜 텍스트 멘션 거부");
            command.Mentions = new[] { new ChatMention(20, new[] { 1, 2 }, 5) };
            Check(!command.TryReadMentionText("/메모", out target, out reason), "복수 멘션 거부");
            command.Mentions = new[] { new ChatMention(20, new[] { 1 }, 5) };
            command.Keyword = "/메모 @이름 공백 " + new string('가', 1001);
            Check(!command.TryReadMentionText("/메모", out target, out reason), "긴 사유 거부");
            command.Keyword = "/메모 @이름 공백 a\nb";
            Check(!command.TryReadMentionText("/메모", out target, out reason), "제어문자 사유 거부");
            Check(OperatorCommandPolicy.RequiresOperator("/메모 @대상") && !OperatorCommandPolicy.RequiresOperator("/메모아님"), "명령어 정확히 구분");
            OperatorCommandPolicy.Register("/공지등록");
            Check(OperatorCommandPolicy.RequiresOperator("/공지등록 내용"), "공통 권한 정책 등록");
            var row = new object[] { 11L, 99L, 0, 110L, "{\"feedType\":15,\"prevHost\":{\"userId\":10},\"newHost\":{\"userId\":20}}", 0, 0, null };
            var changes = ChatMessageDecoder.Decode(new ChatRoomInfo { ChatId = 10, LinkId = 99 }, row, new ChatUserDirectory(), false);
            Check(changes.Count == 2 && changes[0].ExpectedMemberType == -2 && changes[1].ExpectedMemberType == 1 && changes.All(c => c.RoleChanged), "방장 변경 양쪽 권한 갱신");
            Console.WriteLine("PASS: " + count + " operator checks (no network writes or kicks)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
