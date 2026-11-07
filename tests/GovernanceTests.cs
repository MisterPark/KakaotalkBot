using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using KakaotalkBot;
using Newtonsoft.Json;

class GovernanceTests
{
    static long Start = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
    static void Check(bool value, string message) { if (!value) throw new Exception(message); Console.WriteLine("PASS " + message); }
    static string PathForTest() { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cases", Guid.NewGuid().ToString("N"), "governance.json"); }
    static GovernanceRoster Roster(int count, long time)
    { return new GovernanceRoster { At = time, Present = new HashSet<string>(Enumerable.Range(1, count).Select(i => "bot_" + i)) }; }
    static GovernanceStore Open(int count, out string path)
    {
        path = PathForTest(); var store = new GovernanceStore(path); store.Enable(1, true, Start - 100);
        store.Observe(1, Roster(count, Start - 1)); store.Pulse(1, Start);
        store.Observe(1, Roster(count, Start + 1)); return store;
    }
    static JudgeElection Election(GovernanceStore store)
    { var rows = store.ToRows(); return JsonConvert.DeserializeObject<GovernanceState>(string.Concat(rows.Skip(1).Select(r => r[1]))).Rooms[1].Elections.Last(); }
    static void Candidate(GovernanceStore store, int user) { store.Handle(1, "bot_" + user, "후보" + user, "/출마", "공약", Start + 1); }
    static void Vote(GovernanceStore store, int user, string ballot) { store.Handle(1, "bot_" + user, "투표자", "/투표", ballot, Start + 1); }
    static int Main()
    {
        try
        {
            Check(GovernanceStore.NextStart(Start - 1) == Start && GovernanceStore.NextStart(Start) == new DateTimeOffset(2027, 3, 1, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds(), "한국 시간 홀수달 첫날 및 연도 전환");
            string path; var store = Open(83, out path);
            Check(Election(store).Threshold == 9 && Election(store).End == Start + 259200, "83명은 최소 9표, 선거 기간 3일");
            Candidate(store, 1); Vote(store, 1, "1"); Vote(store, 1, "1");
            Check(Election(store).Ballots.Count == 1, "반복 투표는 교체, 중복 집계 없음");
            Vote(store, 2, "1 1"); Check(Election(store).Ballots.Count == 1, "같은 후보 중복 선택 거부");
            store.Observe(1, Roster(84, Start + 2)); Vote(store, 84, "1");
            Check(Election(store).Electors.Count == 83 && Election(store).Ballots.Count == 1, "신규 입장자 투표 제외 및 분모 고정");
            store.Handle(1, "bot_1", "후보1", "/투표취소", "", Start + 2);
            Check(Election(store).Ballots.Count == 0, "투표 철회");
            for (int i = 1; i <= 8; i++) store.Handle(1, "bot_" + i, "유저", "/투표", "1", Start + 2);
            var restored = new GovernanceStore(path);
            Check(Election(restored).Ballots.Count == 8, "재시작 후 투표 복구");
            restored.Pulse(1, Start + 259200);
            Check(Election(restored).Winners.Count == 0 && Election(restored).State == "closed", "10% 미달이면 당선자 0명");
            restored.Pulse(1, Start + 259201); Check(Election(restored).Winners.Count == 0, "마감 재호출 멱등성");
            var exact = Open(83, out path); Candidate(exact, 1);
            for (int i = 1; i <= 9; i++) Vote(exact, i, "1");
            exact.Pulse(1, Start + 259200); exact.Observe(1, Roster(83, Start + 259200));
            Check(Election(exact).Winners.SequenceEqual(new[] { 1 }) && exact.IsJudge(1, "bot_1", Start + 259200), "최소 득표 충족 시 임명");
            long expiry = new DateTimeOffset(2027, 3, 4, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            Check(Election(exact).TermEnd == expiry && !exact.IsJudge(1, "bot_1", expiry), "2개월 임기와 종료 시각");
            exact.Leave(1, "bot_1"); exact.Observe(1, Roster(83, Start + 259201));
            Check(!exact.IsJudge(1, "bot_1", Start + 259201), "퇴장한 대법관은 재입장해도 권한 자동 복구 없음");
            var tie = Open(100, out path); for (int i = 1; i <= 11; i++) Candidate(tie, i);
            for (int i = 1; i <= 11; i++) for (int voter = (i - 1) * 8 + 1; voter <= (i - 1) * 8 + 8; voter++) Vote(tie, voter, i.ToString());
            // 모든 후보에게 10표를 동일하게 부여합니다. 1인 최대 10명 제한을 지킵니다.
            for (int voter = 1; voter <= 11; voter++) Vote(tie, voter, string.Join(" ", Enumerable.Range(1, 11).Where(n => n != voter)));
            // 위의 다른 기존 표는 초기화하고 정확한 동률을 만듭니다.
            for (int voter = 12; voter <= 100; voter++) tie.Handle(1, "bot_" + voter, "", "/투표취소", "", Start + 1);
            tie.Pulse(1, Start + 259200);
            Check(Election(tie).State == "runoff" && Election(tie).Runoff.Count == 11 && Election(tie).Winners.Count == 0, "10위 경계 동률은 24시간 결선");
            tie.Observe(1, Roster(100, Start + 259201)); tie.Handle(1, "bot_1", "", "/투표", "1", Start + 259201);
            tie.Pulse(1, Start + 345600);
            Check(Election(tie).State == "closed" && Election(tie).Winners.SequenceEqual(new[] { 1 }), "결선 종료 및 공석 처리");
            var missing = new GovernanceStore(PathForTest()); missing.Enable(1, true, Start - 50); missing.Observe(1, Roster(20, Start + 100)); missing.Pulse(1, Start + 100);
            Check(missing.PeekNotice(1).Contains("누락"), "시작 명부 없이 소급 선거 금지");
            var backup = exact.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList();
            var otherPc = new GovernanceStore(PathForTest(), backup);
            Check(Election(otherPc).Winners.SequenceEqual(new[] { 1 }), "시트 백업에서 다른 PC 복구");
            backup[1][1] += "x"; bool rejected = false;
            try { new GovernanceStore(PathForTest(), backup); } catch (InvalidDataException) { rejected = true; }
            Check(rejected, "시트 조각 손상 검출");
            exact.RecordBans(1, new[] { new BanRecord { User = "bot_12", Name = "동일닉", KickedAt = Start } });
            exact.RecordBans(1, new[] { new BanRecord { User = "bot_13", Name = "동일닉", KickedAt = Start }, new BanRecord { User = "bot_12", Name = "변경닉", KickedAt = Start } });
            Check(exact.FindBan(1, 1).User == "bot_12" && exact.FindBan(1, 2).User == "bot_13", "해제 번호는 닉네임·목록 순서 변경과 무관하게 ID에 연결");
            Check(exact.QueueModeration(1, "1:9", "kick", "bot_1", "bot_2", "사유", Start) && !exact.QueueModeration(1, "1:9", "kick", "bot_1", "bot_2", "사유", Start), "같은 메시지로 관리 매크로 중복 접수 금지");
            Console.WriteLine("PASS: 선거 테스트 완료 (실제 채팅·시트 쓰기·내보내기 없음)"); return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
