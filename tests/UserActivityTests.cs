using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using KakaotalkBot;

class UserActivityTests
{
    static int passed;
    static void Check(bool condition, string name) { if (!condition) throw new Exception(name); passed++; }
    static List<string> Cells(List<object> row) { return row.Select(v => Convert.ToString(v, CultureInfo.InvariantCulture)).ToList(); }
    static int Main()
    {
        try
        {
            var user = new User(9007199254740993L);
            foreach (var pair in new[] { Tuple.Create(0L, 1), Tuple.Create(99L, 1), Tuple.Create(100L, 2), Tuple.Create(299L, 2), Tuple.Create(300L, 3), Tuple.Create(600L, 4) })
            { user.Experience = pair.Item1; Check(user.Level == pair.Item2, "레벨 경계 " + pair.Item1); }
            user.Experience = long.MaxValue;
            Check(user.NextLevelExperience > user.Experience, "큰 경험치의 다음 레벨 계산");
            Check(User.ToUser(Cells(user.ToRow())).Experience == long.MaxValue, "경험치 정밀도 보존");
            Check(User.ToUser(Cells(user.ToRow()).Take(10).ToList()).Level == 1, "기존 사용자 기본 레벨");
            user.Experience = 0;
            var store = new UserActivityStore();
            Check(store.AwardExperience(user, 10, 1, 100), "첫 채팅 경험치");
            Check(!store.AwardExperience(user, 10, 2, 159), "59초 도배 제한");
            Check(!store.AwardExperience(user, 20, 3, 159), "다른 방에서도 사용자별 제한");
            Check(store.AwardExperience(user, 20, 4, 160), "60초 지급");
            Check(!store.AwardExperience(user, 20, 4, 220), "중복 로그 지급 방지");
            Check(!store.AwardExperience(user, 20, 5, 150), "과거 시각 지급 방지");
            Check(user.Experience == 2, "총 경험치");
            store.Observe(10, user.UserId, "이름1", true, 100, 10, "profile");
            Check(store.Nicknames.Count == 0, "최초 이름을 변경 이력으로 만들지 않음");
            store.Observe(10, user.UserId, "이름2", true, 200, 20, "profile");
            store.Observe(10, user.UserId, "이름2", true, 201, 20, "profile");
            Check(store.Nicknames.Count == 1 && store.Nicknames[0].Before == "이름1", "프로필 변경 이력과 반복 관찰 중복 제외");
            store.Observe(20, user.UserId, "다른 방 이름", true, 200, 20, "profile");
            Check(store.Nicknames.Count == 1, "방별 닉네임 분리");
            Check(store.RecordEvent(10, user.UserId, 30, "kick", "이름2", 300), "강퇴 이력");
            Check(store.Get(10, user.UserId).IsPresent == false, "강퇴 상태");
            Check(store.RecordEvent(10, user.UserId, 40, "join", "이름3", 400), "재입장 이력");
            Check(store.RecordEvent(10, user.UserId, 35, "leave", "이름2", 350), "늦은 퇴장 이력도 보존");
            Check(store.Get(10, user.UserId).IsPresent == true && store.Get(10, user.UserId).Nickname == "이름3", "늦은 퇴장이 재입장 상태를 덮어쓰지 않음");
            store.Observe(10, user.UserId, "오래된 프로필", false, 500, 36, "profile");
            Check(store.Get(10, user.UserId).IsPresent == true && store.Get(10, user.UserId).Nickname == "이름3", "지연된 참여자 스냅샷 제외");
            Check(!store.RecordEvent(10, user.UserId, 40, "join", "이름3", 400), "중복 입장 이력 제외");
            var restored = new UserActivityStore();
            restored.Load(store.Rooms.Values.Select(r => Cells(r.ToRow())).ToList(), store.Events.Select(r => Cells(r.ToRow())).ToList(), store.Nicknames.Select(r => Cells(r.ToRow())).ToList());
            Check(restored.SavedEvents == 3 && restored.SavedNicknames == 2 && !restored.Dirty, "이력 재시작 로드와 저장 위치");
            Check(!restored.RecordEvent(10, user.UserId, 30, "kick", "이름2", 300), "재시작 후 강퇴 중복 방지");
            Check(!restored.AwardExperience(user, 20, 6, 180), "재시작 후 쿨다운 유지");
            Check(restored.Rooms[Tuple.Create(10L, user.UserId)].IsPresent == true, "방 상태 저장 왕복");
            Console.WriteLine("PASS: " + passed + " activity checks (no network writes or chat sends)");
            return 0;
        }
        catch(Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
