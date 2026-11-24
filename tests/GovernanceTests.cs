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
            var disabled = new GovernanceStore(PathForTest());
            Check(disabled.QueueModeration(1, "1:10", "/내보낸목록", "bot_1", "", "", Start) && !disabled.Enabled(1), "선거를 켜지 않고 관리 명령 최초 접수");
            disabled.RecordBans(1, new[] { new BanRecord { User = "bot_2", Name = "대상", KickedAt = Start } });
            disabled.FinishModeration(1, "1:10", null);
            Check(disabled.PeekNotice(1).Contains("대상") && disabled.FindBan(1, 1) != null, "선거 미활성 방에서도 목록 결과 저장·응답");
            disabled.Enable(1, true, Start);
            Check(disabled.Enabled(1) && disabled.FindBan(1, 1) != null, "관리 이력을 보존하며 선거 활성화");
            disabled.Enable(1, false, Start);
            Check(disabled.QueueModeration(1, "1:11", "/내보낸목록", "bot_1", "", "", Start), "선거를 끈 뒤에도 관리 명령 접수");
            while (disabled.PeekNotice(1) != null) disabled.TakeNotice(1);
            disabled.RecordBans(1, new BanRecord[0]);
            Check(disabled.PeekNotice(1).Contains("내보낸 상대가 없습니다.") && !disabled.PeekNotice(1).Contains("/내보내기해제") && disabled.FindBan(1, 1) == null,
                "빈 목록을 정상 응답하고 이전 번호를 해제 대상으로 사용하지 않음");
            long november = new DateTimeOffset(2026, 11, 11, 12, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            var manual = new GovernanceStore(PathForTest()); manual.Enable(1, true, november); manual.Observe(1, Roster(100, november));
            Check(manual.OpenElection(1, false, november).Contains("제1회") && Election(manual).Start == november && Election(manual).End == november + 3 * 86400,
                "놓친 11월 정기선거를 실제 개시 시점부터 3일간 진행");
            long january = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            Check(Election(manual).TermEnd == january + 3 * 86400 && manual.Handle(1, "bot_1", "", "/선거", "", november).Contains("2027-01-01"),
                "늦게 연 선거도 다음 정기 일정과 원래 임기 종료 유지");
            Check(manual.OpenElection(1, false, november + 1).Contains("진행 중") && manual.OpenElection(1, true, november + 1).Contains("진행 중"), "본선 진행 중 중복 선거 금지");
            for (int i = 1; i <= 10; i++) manual.Handle(1, "bot_" + i, "후보" + i, "/출마", "공약", november + 1);
            for (int i = 1; i <= 10; i++) manual.Handle(1, "bot_" + i, "", "/투표", "1 2 3 4 5 6 7 8 9 10", november + 1);
            long byStart = november + 3 * 86400 + 1;
            manual.Pulse(1, byStart); manual.Observe(1, Roster(100, byStart));
            Check(Election(manual).Winners.Count == 10 && manual.OpenElection(1, true, byStart).Contains("정원 10명"), "전원 재임 중 보궐선거 금지");
            Check(manual.OpenElection(1, false, byStart).Contains("이미 개최"), "같은 회차 정기선거 재개최로 당선자를 늘리지 않음");
            var missedLeave = new GovernanceStore(PathForTest(), manual.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList());
            var withoutTenth = Roster(100, byStart); withoutTenth.Present.Remove("bot_10"); missedLeave.Observe(1, withoutTenth);
            Check(missedLeave.OpenElection(1, true, byStart).Contains("보궐선거") && Election(missedLeave).Seats == 1, "퇴장 이벤트 누락도 최신 명부로 공석 확인");
            missedLeave.Observe(1, Roster(100, byStart));
            Check(!missedLeave.IsJudge(1, "bot_10", byStart), "공석 보충 이후 복귀한 전임자의 중복 임명 방지");
            manual.Leave(1, "bot_10"); manual.Observe(1, Roster(100, byStart));
            Check(manual.OpenElection(1, true, byStart).Contains("제2회") && Election(manual).Seats == 1 && Election(manual).Kind == "by", "퇴장한 대법관의 빈자리 1석만 보충");
            Check(manual.IsJudge(1, "bot_1", byStart) && !manual.IsJudge(1, "bot_10", byStart), "보궐선거 중 기존 임기 유지와 퇴장자 복귀 권한 금지");
            Check(manual.Handle(1, "bot_1", "후보1", "/출마", "공약", byStart).Contains("재임 중"), "현직 대법관 보궐 중복 출마 금지");
            manual.Handle(1, "bot_11", "후보11", "/출마", "공약", byStart);
            manual.Handle(1, "bot_12", "후보12", "/출마", "공약", byStart);
            Check(manual.Handle(1, "bot_1", "", "/투표", "1 2", byStart).Contains("선택 인원"), "보궐선거 승인투표 인원을 공석 수로 제한");
            for (int i = 1; i <= 20; i++) manual.Handle(1, "bot_" + i, "", "/투표", i <= 10 ? "1" : "2", byStart);
            long runoffAt = byStart + 3 * 86400;
            manual.Pulse(1, runoffAt); manual.Observe(1, Roster(100, runoffAt));
            Check(Election(manual).State == "runoff" && manual.OpenElection(1, true, runoffAt).Contains("진행 중"), "1석 보궐 동률 결선과 결선 중 중복 개시 금지");
            manual.Handle(1, "bot_1", "", "/투표", "1", runoffAt);
            long byEnd = runoffAt + 86400;
            manual.Pulse(1, byEnd); manual.Observe(1, Roster(100, byEnd));
            Check(Election(manual).Winners.SequenceEqual(new[] { 1 }) && Election(manual).TermEnd == january + 3 * 86400 &&
                Enumerable.Range(1, 100).Count(i => manual.IsJudge(1, "bot_" + i, byEnd)) == 10, "보궐 당선자는 잔여 임기만 맡고 전체 정원은 10명 유지");
            string history = manual.Handle(1, "bot_1", "", "/선거기록", "", byEnd);
            Check(history.Contains("제1회 정기선거") && history.Contains("제2회 보궐선거") && history.Contains("후보11 · 본선 10표 · 결선 1표 · 당선") && history.Contains("후보12 · 본선 10표 · 결선 0표 · 미당선"), "회차별 본선·결선 득표수와 당선 여부 모두 표시");
            Check(!manual.History(1, "2").Contains("제1회") && manual.History(1, "3").Contains("형식"), "선거기록 특정 회차 조회와 잘못된 번호 안내");
            manual.Enable(1, false, byEnd);
            Check(manual.History(1, "").Contains("후보11") && GovernanceStore.Messages(history).All(p => p.Length < 3100), "선거를 꺼도 기록 조회 가능하며 긴 기록 분할");
            var manualBackup = manual.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList();
            Check(new GovernanceStore(PathForTest(), manualBackup).History(1, "") == history, "선거 회차·득표·정원·종류가 다른 PC에서도 복구");
            var longHistory = string.Join("\n", Enumerable.Range(1, 300).Select(i => "후보" + i + new string('가', 30) + " · 본선 10표 · 당선"));
            var historyPages = GovernanceStore.Messages(longHistory).ToArray();
            Check(historyPages.Length > 1 && historyPages.All(p => p.Length < 3100) && Enumerable.Range(1, 300).All(i => historyPages.Any(p => p.Contains("후보" + i + new string('가', 30)))), "긴 선거기록은 모든 후보를 보존하여 분할");
            var stale = new GovernanceStore(PathForTest()); stale.Enable(1, true, november);
            Check(stale.OpenElection(1, false, november).Contains("명부"), "명부 없는 강제 개시 금지");
            stale.Observe(1, Roster(100, november - 31));
            Check(stale.OpenElection(1, false, november).Contains("명부"), "오래된 선거인 명부로 강제 개시 금지");
            stale.Observe(1, Roster(100, january - 86400));
            Check(stale.OpenElection(1, false, january - 86400).Contains("겹쳐"), "다음 정기선거와 겹치는 강제 개시 금지");
            Console.WriteLine("PASS: 선거 테스트 완료 (실제 채팅·시트 쓰기·내보내기 없음)"); return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
