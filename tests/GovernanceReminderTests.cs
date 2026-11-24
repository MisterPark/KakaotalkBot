using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using KakaotalkBot;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

internal static class GovernanceReminderTests
{
    private static readonly long Start = new DateTimeOffset(2027, 1, 1, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
    private static int passed;
    private static void Check(bool value, string label) { if (!value) throw new Exception(label); passed++; }
    private static string NewPath() { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cases", Guid.NewGuid().ToString("N"), "state.json"); }
    private static GovernanceState State(GovernanceStore store)
    { return JsonConvert.DeserializeObject<GovernanceState>(string.Concat(store.ToRows().Skip(1).Select(r => r[1]))); }
    private static JudgeElection Election(GovernanceStore store) { return State(store).Rooms[1].Elections.Last(); }
    private static string[] Reminders(GovernanceStore store) { return State(store).Rooms[1].Notices.Where(GovernanceStore.IsVotingReminder).ToArray(); }
    private static void Drain(GovernanceStore store) { while (store.PeekNotice(1) != null) store.TakeNotice(1); }
    private static GovernanceRoster Roster(long at)
    { return new GovernanceRoster { At = at, Present = new HashSet<string>(Enumerable.Range(1, 20).Select(i => "bot_" + i)) }; }
    private static GovernanceStore Open(out string path)
    {
        path = NewPath(); var store = new GovernanceStore(path);
        store.Enable(1, true, Start - 1); store.Observe(1, Roster(Start - 1)); store.Pulse(1, Start);
        store.Observe(1, Roster(Start + 1)); Drain(store); return store;
    }
    private static GovernanceStore Restore(GovernanceState state)
    {
        string path = NewPath(); Directory.CreateDirectory(Path.GetDirectoryName(path));
        File.WriteAllText(path, JsonConvert.SerializeObject(state)); return new GovernanceStore(path);
    }
    private static int Main()
    {
        try
        {
            string path; var store = Open(out path);
            store.Handle(1, "bot_1", "첫 후보", "/출마", "공약", Start + 1);
            store.Handle(1, "bot_2", "둘째 후보", "/출마", "공약", Start + 1);
            store.Handle(1, "bot_3", "유권자", "/투표", "1", Start + 1);
            Check(Election(store).NextReminderAt == Start + 3600, "First reminder is an hour after opening");
            store.Pulse(1, Start + 3599); Check(store.PeekNotice(1) == null, "No early reminder");
            store.Pulse(1, Start + 3600);
            string first = store.PeekNotice(1);
            Check(Reminders(store).Length == 1 && first.Contains("제1회 정기선거 · 본선투표"), "Regular election reminder at hour boundary");
            Check(first.Contains("2027-01-04 00:00 (한국 시간)") && first.Contains("1/20명") && first.Contains("예: /투표 1 2") && first.Contains("최대 10명"), "Accurate deadline, turnout and voting instructions");
            Check(store.CanSendNotice(1, first, Start + 3600), "Active reminder allowed");
            store.Pulse(1, Start + 3600); store.Pulse(1, Start + 3601);
            Check(Reminders(store).Length == 1, "Repeated ticks do not duplicate reminders");
            store = new GovernanceStore(path); store.Pulse(1, Start + 3602);
            Check(Reminders(store).Length == 1 && Election(store).NextReminderAt == Start + 7200, "Pending reminder and next time survive restart");
            Drain(store); store = new GovernanceStore(path); store.Pulse(1, Start + 3603);
            Check(store.PeekNotice(1) == null, "Consumed reminder is not recreated after restart");
            store.Pulse(1, Start + 7200);
            Check(Reminders(store).Length == 1 && Election(store).NextReminderAt == Start + 10800, "Second hourly reminder");
            long resumed = Start + 10 * 3600 + 17;
            store.Pulse(1, resumed);
            Check(Reminders(store).Length == 1 && Election(store).NextReminderAt == resumed + 3600, "Downtime produces one latest reminder without catch-up burst");
            var backup = store.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList();
            var anotherPc = new GovernanceStore(NewPath(), backup); anotherPc.Pulse(1, resumed + 1);
            Check(Reminders(anotherPc).Length == 1 && Election(anotherPc).NextReminderAt == resumed + 3600, "Sheet restore preserves reminder schedule");
            Drain(store); store.Pulse(1, Start + 1);
            Check(store.PeekNotice(1) == null, "Clock rollback does not repeat a reminder");
            store.Pulse(1, resumed + 3600);
            store.Pulse(1, Start + 3 * 86400);
            Check(Reminders(store).Length == 0 && Election(store).State == "closed", "Deadline removes pending reminders and stops schedule");
            Check(!store.CanSendNotice(1, first, Start + 3 * 86400), "Queued reminder rejected at deadline");
            Drain(store); store.Pulse(1, Start + 3 * 86400 + 3600);
            Check(store.PeekNotice(1) == null, "No post-election reminder");

            var disabled = Open(out path); disabled.Pulse(1, Start + 3600);
            string pending = disabled.PeekNotice(1);
            Check(pending.Contains("/출마 공약") && !pending.Contains("예: /투표"), "No-candidate reminder invites registration without invalid candidate numbers");
            disabled.RecordBans(1, new BanRecord[0]); disabled.Enable(1, false, Start + 3601);
            disabled.Pulse(1, Start + 7200);
            Check(Reminders(disabled).Length == 0 && disabled.PeekNotice(1).StartsWith("[내보낸 목록]"), "Disabling removes reminders but preserves unrelated notices");
            Check(!disabled.CanSendNotice(1, pending, Start + 3602) && disabled.CanSendNotice(1, "일반 안내", Start + 3602), "Send-time disable guard only suppresses reminders");

            long manualStart = new DateTimeOffset(2026, 11, 11, 12, 34, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            var manual = new GovernanceStore(NewPath()); manual.Enable(1, true, manualStart); manual.Observe(1, Roster(manualStart));
            manual.OpenElection(1, false, manualStart); manual.Pulse(1, manualStart + 3599);
            Check(manual.PeekNotice(1) == null && Election(manual).NextReminderAt == manualStart + 3600, "Forced election keeps its own hourly start time");
            manual.Pulse(1, manualStart + 3600); Check(Reminders(manual).Length == 1, "Forced election reminders enabled");

            var source = Open(out path); var state = State(source); var election = state.Rooms[1].Elections[0];
            election.Kind = "by"; election.Seats = 1;
            election.Candidates.Add(new ElectionCandidate { Number = 1, User = "bot_1", Name = "후보1", Pledge = "공약" });
            election.Candidates.Add(new ElectionCandidate { Number = 2, User = "bot_2", Name = "후보2", Pledge = "공약" });
            election.Ballots.Add("bot_1", new List<int> { 1 }); election.Ballots.Add("bot_2", new List<int> { 1 });
            election.Ballots.Add("bot_3", new List<int> { 2 }); election.Ballots.Add("bot_4", new List<int> { 2 });
            var by = Restore(state); by.Pulse(1, Start + 3600); string byNotice = by.PeekNotice(1);
            Check(byNotice.Contains("보궐선거") && byNotice.Contains("최대 1명") && byNotice.Contains("예: /투표 1\n"), "By-election reminder respects vacancy limit");
            long runoffStart = Start + 3 * 86400;
            Check(!by.CanSendNotice(1, byNotice, runoffStart), "Main ballot notice expires even before closing tick");
            by.Pulse(1, runoffStart);
            Check(Election(by).State == "runoff" && Reminders(by).Length == 0, "Runoff transition removes main reminder without an immediate extra reminder");
            Check(!by.CanSendNotice(1, byNotice, runoffStart + 1), "Main ballot notice stays invalid during runoff");
            Drain(by); by.Pulse(1, runoffStart + 3599); Check(by.PeekNotice(1) == null, "Runoff waits a full hour");
            by.Pulse(1, runoffStart + 3600); string runoffNotice = by.PeekNotice(1);
            Check(runoffNotice.Contains("결선투표") && runoffNotice.Contains("0/20명") && runoffNotice.Contains("2027-01-05 00:00") && runoffNotice.Contains("최대 1명"), "Runoff reminder uses runoff ballots, deadline and remaining seats");
            Check(by.CanSendNotice(1, runoffNotice, runoffStart + 3601), "Current runoff reminder is allowed");
            by.Pulse(1, runoffStart + 86400);
            Check(Reminders(by).Length == 0 && !by.CanSendNotice(1, runoffNotice, runoffStart + 86400), "Runoff deadline suppresses queued reminders");
            var delayed = Restore(state); delayed.Pulse(1, runoffStart + 7200);
            Check(Election(delayed).State == "runoff" && Reminders(delayed).Length == 0 && Election(delayed).NextReminderAt == runoffStart + 10800, "Delayed runoff announcement starts a new hour instead of sending two notices");

            source = Open(out path);
            var legacy = JObject.Parse(File.ReadAllText(path));
            foreach (JObject room in ((JObject)legacy["Rooms"]).Properties().Select(p => p.Value))
                foreach (JObject oldElection in (JArray)room["Elections"]) oldElection.Remove("NextReminderAt");
            File.WriteAllText(path, legacy.ToString());
            var migrated = new GovernanceStore(path); migrated.Pulse(1, Start + 2 * 3600);
            Check(Reminders(migrated).Length == 1 && Election(migrated).NextReminderAt == Start + 3 * 3600, "Existing elections without reminder fields migrate safely");
            Check(!migrated.CanSendNotice(2, migrated.PeekNotice(1), Start + 7200), "Reminder cannot be sent for a different room");
            Console.WriteLine("PASS: " + passed + " election reminder checks (synthetic state only; no chats or sheet writes)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
