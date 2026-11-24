using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    // 사용자 키는 BotIdentity의 bot_ ID만 사용합니다. 카카오 ID는 저장하지 않습니다.
    internal sealed class ElectionCandidate
    {
        public int Number;
        public string User, Name, Pledge;
        public bool Withdrawn;
    }
    internal sealed class JudgeElection
    {
        public long Start, End, TermEnd;
        public long CycleStart;
        public long NextReminderAt;
        public string Kind = "regular";
        public int Seats = 10;
        public string State = "voting";
        public List<string> Electors = new List<string>();
        public List<ElectionCandidate> Candidates = new List<ElectionCandidate>();
        public Dictionary<string, List<int>> Ballots = new Dictionary<string, List<int>>();
        public List<int> Winners = new List<int>(), Runoff = new List<int>();
        public Dictionary<string, List<int>> RunoffBallots = new Dictionary<string, List<int>>();
        public int Threshold { get { return (Electors.Count + 9) / 10; } }
    }
    internal sealed class BanRecord
    {
        public int Number;
        public string User, Name;
        public long KickedAt;
        public bool Listed;
    }
    internal sealed class ModerationRecord
    {
        public string Key, Kind, Actor, Target, Reason, Status, Error;
        public long At;
    }
    internal sealed class JudgeRoom
    {
        public long ChatId, NextElection;
        public bool Enabled;
        public List<JudgeElection> Elections = new List<JudgeElection>();
        public List<string> Notices = new List<string>();
        public HashSet<string> Revoked = new HashSet<string>();
        public List<BanRecord> Bans = new List<BanRecord>();
        public List<ModerationRecord> Moderation = new List<ModerationRecord>();
    }
    internal sealed class GovernanceState
    {
        public int Version = 1;
        public long Revision;
        public Dictionary<long, JudgeRoom> Rooms = new Dictionary<long, JudgeRoom>();
    }
    internal sealed class GovernanceRoster
    {
        public long At;
        public HashSet<string> Present = new HashSet<string>();
        public Dictionary<string, string> Names = new Dictionary<string, string>();
    }
    internal sealed class GovernanceStore
    {
        internal static readonly string[] Headers = { "part", "payload" };
        private GovernanceState state;
        private readonly string path;
        private readonly Dictionary<long, GovernanceRoster> live = new Dictionary<long, GovernanceRoster>();
        internal bool Dirty { get; private set; }
        internal GovernanceStore(string path, List<List<string>> sheet = null)
        {
            this.path = path;
            var local = File.Exists(path) ? Parse(File.ReadAllText(path)) : new GovernanceState();
            var remote = ReadSheet(sheet);
            state = remote != null && remote.Revision > local.Revision ? remote : local;
            Dirty = remote == null || state.Revision > remote.Revision;
            if (!ReferenceEquals(state, local)) SaveFile(state);
        }
        private static GovernanceState Parse(string json)
        {
            var value = JsonConvert.DeserializeObject<GovernanceState>(json);
            if (value == null || value.Version != 1 || value.Rooms == null) throw new InvalidDataException("선거 저장 형식 오류");
            foreach (var room in value.Rooms.Values)
                foreach (var election in room.Elections)
                {
                    if (election.CycleStart == 0) election.CycleStart = CycleStart(election.Start);
                    // 기존 진행 중 선거도 새 알림 주기에 참여합니다. 지난 시간은 Pulse에서 한 번만 처리합니다.
                    if (election.NextReminderAt == 0 && (election.State == "voting" || election.State == "runoff"))
                        election.NextReminderAt = (election.State == "runoff" ? election.End : election.Start) + 3600;
                    if (election.Seats < 1 || election.Seats > 10 || (election.Kind != "regular" && election.Kind != "by"))
                        throw new InvalidDataException("선거 종류 또는 선출 정원 오류");
                    if (election.Electors.Any(x => x == null || !x.StartsWith("bot_", StringComparison.Ordinal)) || election.Electors.Count != election.Electors.Distinct().Count())
                        throw new InvalidDataException("선거인 명부의 고정 ID 형식 오류");
                }
            return value;
        }
        private static string Hash(string json)
        { using (var hash = SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(json))); }
        private static GovernanceState ReadSheet(List<List<string>> rows)
        {
            if (rows == null || rows.Count == 0) return null;
            int count;
            if (rows[0].Count != 2 || !rows[0][0].StartsWith("manifest:") || !int.TryParse(rows[0][0].Substring(9), out count) || count < 1 || count > rows.Count - 1)
                throw new InvalidDataException("선거 시트 매니페스트 오류");
            var text = new StringBuilder();
            for (int i = 0; i < count; i++)
            { if (rows[i + 1].Count != 2 || rows[i + 1][0] != "part:" + i) throw new InvalidDataException("선거 시트 조각 누락"); text.Append(rows[i + 1][1]); }
            if (Hash(text.ToString()) != rows[0][1]) throw new InvalidDataException("선거 시트 검증 실패");
            return Parse(text.ToString());
        }
        internal List<List<object>> ToRows()
        {
            string json = JsonConvert.SerializeObject(state);
            var rows = new List<List<object>> { new List<object> { "manifest:" + ((json.Length + 23999) / 24000), Hash(json) } };
            for (int i = 0; i < json.Length; i += 24000) rows.Add(new List<object> { "part:" + i / 24000, json.Substring(i, Math.Min(24000, json.Length - i)) });
            return rows;
        }
        internal void MarkSaved() { Dirty = false; }
        private void SaveFile(GovernanceState value)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            File.WriteAllText(temp, JsonConvert.SerializeObject(value), new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, path + ".bak"); else File.Move(temp, path);
        }
        private void Commit(GovernanceState next)
        {
            next.Revision = Math.Max(state.Revision + 1, DateTime.UtcNow.Ticks);
            SaveFile(next); state = next; Dirty = true;
        }
        private GovernanceState Copy() { return Parse(JsonConvert.SerializeObject(state)); }
        internal static DateTimeOffset Local(long now) { return DateTimeOffset.FromUnixTimeSeconds(now).ToOffset(TimeSpan.FromHours(9)); }
        internal static long NextStart(long now)
        {
            var local = Local(now);
            var first = new DateTimeOffset(local.Year, local.Month, 1, 0, 0, 0, TimeSpan.FromHours(9));
            while (first.Month % 2 == 0 || first.ToUnixTimeSeconds() <= now) first = first.AddMonths(1);
            return first.ToUnixTimeSeconds();
        }
        internal static long CycleStart(long now)
        {
            var local = Local(now);
            var first = new DateTimeOffset(local.Year, local.Month, 1, 0, 0, 0, TimeSpan.FromHours(9));
            return (first.Month % 2 == 0 ? first.AddMonths(-1) : first).ToUnixTimeSeconds();
        }
        private static JudgeElection NewElection(long start, long cycle, GovernanceRoster roster, int seats, bool byElection)
        {
            return new JudgeElection { Start = start, End = start + 3 * 86400, CycleStart = cycle, NextReminderAt = start + 3600,
                TermEnd = Local(cycle).AddMonths(2).AddDays(3).ToUnixTimeSeconds(), Seats = seats, Kind = byElection ? "by" : "regular",
                Electors = roster.Present.OrderBy(x => x, StringComparer.Ordinal).ToList() };
        }
        private static string OpeningNotice(JudgeElection election, int number)
        {
            return "[제" + number + "회 대법관 " + (election.Kind == "by" ? "보궐선거" : "정기선거") + " 시작]\n선출 " + election.Seats + "명 · 선거인 " + election.Electors.Count + "명 · 당선 기준 " + election.Threshold + "표\n" +
                Local(election.End).ToString("yyyy-MM-dd HH:mm") + " 마감\n임기 종료: " + Local(election.TermEnd).ToString("yyyy-MM-dd HH:mm") +
                "\n/출마 공약 · /후보 · /투표 후보번호 (최대 " + election.Seats + "명, 공개투표)";
        }
        private const string ReminderTitle = "[대법관 투표 참여 안내]";
        internal static bool IsVotingReminder(string notice)
        { return notice != null && notice.StartsWith(ReminderTitle + "\n", StringComparison.Ordinal); }
        private static bool IsVoting(JudgeElection election, long now)
        {
            return now >= election.Start && (election.State == "voting" && now < election.End ||
                election.State == "runoff" && now >= election.End && now < election.End + 86400);
        }
        private static string ReminderHeading(JudgeElection election, int number)
        {
            return ReminderTitle + "\n제" + number + "회 " + (election.Kind == "by" ? "보궐선거" : "정기선거") +
                " · " + (election.State == "runoff" ? "결선투표" : "본선투표") + "\n";
        }
        private static string VotingReminder(JudgeElection election, int number)
        {
            bool runoff = election.State == "runoff";
            int limit = election.Seats - (runoff ? election.Winners.Count : 0);
            var candidates = election.Candidates.Where(c => !c.Withdrawn && (!runoff || election.Runoff.Contains(c.Number))).OrderBy(c => c.Number).ToArray();
            string example = string.Join(" ", candidates.Take(Math.Min(2, limit)).Select(c => c.Number));
            return ReminderHeading(election, number) + "아직 투표하지 않았다면 참여해 주세요!\n" +
                "마감: " + Local(election.End + (runoff ? 86400 : 0)).ToString("yyyy-MM-dd HH:mm") + " (한국 시간)\n" +
                "투표 참여: " + (runoff ? election.RunoffBallots.Count : election.Ballots.Count) + "/" + election.Electors.Count + "명\n" +
                "/후보 — 후보 번호·공약 확인\n" +
                (candidates.Length == 0 ? (runoff ? "현재 투표할 후보가 없습니다. /선거로 진행 상황을 확인하세요." : "등록된 후보가 없습니다. /출마 공약으로 후보를 등록해 주세요.") :
                    "/투표 후보번호 — 예: /투표 " + example + "\n최대 " + limit + "명 선택 · 공개투표") +
                "\n선거 시작 당시 참여자만 투표할 수 있습니다.";
        }
        internal bool CanSendNotice(long chat, string notice, long now)
        {
            if (!IsVotingReminder(notice)) return true;
            JudgeRoom room;
            if (!state.Rooms.TryGetValue(chat, out room) || !room.Enabled) return false;
            return room.Elections.Select((e, index) => new { Election = e, Number = index + 1 }).Any(item =>
                IsVoting(item.Election, now) && notice.StartsWith(ReminderHeading(item.Election, item.Number), StringComparison.Ordinal));
        }
        private static IEnumerable<string> Mandates(JudgeRoom room, long now)
        {
            return room.Elections.Where(e => now >= e.End && now < e.TermEnd)
                .SelectMany(e => e.Candidates.Where(c => e.Winners.Contains(c.Number) && !room.Revoked.Contains(e.Start + ":" + c.User)).Select(c => c.User)).Distinct();
        }
        internal string OpenElection(long chat, bool byElection, long now)
        {
            Pulse(chat, now);
            JudgeRoom current; GovernanceRoster roster;
            if (!state.Rooms.TryGetValue(chat, out current) || !current.Enabled) return "먼저 /선거설정 켜기로 선거를 활성화하세요.";
            if (current.Elections.Any(e => e.State != "closed")) return "진행 중인 선거 또는 결선이 있어 새 선거를 열 수 없습니다.";
            if (!live.TryGetValue(chat, out roster) || now < roster.At || now - roster.At > 30 || roster.Present.Count == 0)
                return "최신 선거인 명부를 확인하지 못했습니다. 수신 상태를 확인한 뒤 다시 시도하세요.";
            long cycle = CycleStart(now);
            bool regularExists = current.Elections.Any(e => e.Kind == "regular" && e.CycleStart == cycle);
            if (!byElection && regularExists) return "이번 회차의 정기선거는 이미 개최했습니다. 빈자리는 /보궐선거로 선출하세요.";
            if (byElection && !regularExists) return "이번 회차의 정기선거가 없습니다. 먼저 /선거시작으로 정기선거를 여세요.";
            if (now + 4 * 86400 > current.NextElection) return "투표·결선 기간이 다음 정기선거와 겹쳐 지금은 개시할 수 없습니다.";
            var next = Copy(); var room = next.Rooms[chat];
            // 퇴장 이벤트를 놓쳤더라도 공석을 보충한 뒤 복귀한 전임자가 중복 임명되지 않게 합니다.
            foreach (var term in room.Elections.Where(e => now >= e.End && now < e.TermEnd))
                foreach (var candidate in term.Candidates.Where(c => term.Winners.Contains(c.Number) && !roster.Present.Contains(c.User)))
                    room.Revoked.Add(term.Start + ":" + candidate.User);
            int seats = byElection ? 10 - Mandates(room, now).Count() : 10;
            if (seats <= 0) return "현재 대법관 정원 10명이 모두 재임 중이어서 보궐선거를 열 수 없습니다.";
            var election = NewElection(now, cycle, roster, seats, byElection);
            room.Notices.RemoveAll(n => n.StartsWith("[대법관 선거 일정 누락]", StringComparison.Ordinal));
            room.Elections.Add(election); Commit(next);
            return OpeningNotice(election, room.Elections.Count);
        }
        internal void Enable(long chat, bool enabled, long now)
        {
            var next = Copy(); JudgeRoom room;
            if (!next.Rooms.TryGetValue(chat, out room)) next.Rooms.Add(chat, room = new JudgeRoom { ChatId = chat, NextElection = NextStart(now) });
            if (enabled && !room.Enabled) room.NextElection = NextStart(now);
            room.Enabled = enabled;
            if (!enabled) room.Notices.RemoveAll(IsVotingReminder);
            Commit(next);
        }
        internal bool Enabled(long chat) { JudgeRoom room; return state.Rooms.TryGetValue(chat, out room) && room.Enabled; }
        internal void Observe(long chat, GovernanceRoster roster) { live[chat] = roster; }
        internal void Pulse(long chat, long now)
        {
            JudgeRoom current;
            if (!state.Rooms.TryGetValue(chat, out current) || !current.Enabled) return;
            bool due = now >= current.NextElection || current.Elections.Any(e => e.State == "voting" && now >= e.End || e.State == "runoff" && now >= e.End + 86400 ||
                IsVoting(e, now) && now >= e.NextReminderAt);
            if (!due) return;
            var next = Copy(); var room = next.Rooms[chat];
            foreach (var election in room.Elections)
            {
                if (election.State == "voting" && now >= election.End) Close(election, room, false, now);
                if (election.State == "runoff" && now >= election.End + 86400) Close(election, room, true, now);
            }
            if (now >= room.NextElection)
            {
                long start = room.NextElection;
                GovernanceRoster roster;
                bool available = !room.Elections.Any(e => e.State != "closed") && now - start <= 30 && live.TryGetValue(chat, out roster) && roster.At <= start && start - roster.At <= 30;
                if (available)
                {
                    roster = live[chat];
                    var election = NewElection(start, CycleStart(start), roster, 10, false);
                    room.Elections.Add(election);
                    room.Notices.Add(OpeningNotice(election, room.Elections.Count));
                }
                else room.Notices.Add("[대법관 선거 일정 누락]\n선거 시작 당시 참여자 명부를 확보하지 못했거나 다른 선거가 진행 중이라 이번 선거를 열지 않았습니다.\n운영자는 /선거시작으로 놓친 정기선거를 열 수 있습니다.");
                room.NextElection = NextStart(now);
            }
            for (int i = 0; i < room.Elections.Count; i++)
            {
                var election = room.Elections[i];
                if (!IsVoting(election, now) || now < election.NextReminderAt) continue;
                // 대기 중 알림은 최신 한 건으로 교체하고, 누락된 시간별 알림은 소급 발송하지 않습니다.
                room.Notices.RemoveAll(IsVotingReminder);
                room.Notices.Add(VotingReminder(election, i + 1));
                election.NextReminderAt = now + 3600;
            }
            Commit(next);
        }
        private static void Close(JudgeElection e, JudgeRoom room, bool runoff, long now)
        {
            room.Notices.RemoveAll(IsVotingReminder);
            var candidates = e.Candidates.Where(c => !c.Withdrawn && (!runoff || e.Runoff.Contains(c.Number))).ToArray();
            var ballots = runoff ? e.RunoffBallots : e.Ballots;
            var scores = candidates.Select(c => new { c.Number, Count = ballots.Values.Count(v => v.Contains(c.Number)) }).Where(x => runoff ? x.Count > 0 : x.Count >= e.Threshold && x.Count > 0).OrderByDescending(x => x.Count).ThenBy(x => x.Number).ToArray();
            int seats = e.Seats - e.Winners.Count;
            foreach (var group in scores.GroupBy(x => x.Count))
            {
                if (group.Count() <= seats) { e.Winners.AddRange(group.Select(x => x.Number)); seats -= group.Count(); }
                else if (seats > 0 && !runoff)
                { e.Runoff = group.Select(x => x.Number).ToList(); e.State = "runoff"; e.NextReminderAt = now + 3600; room.Notices.Add("[대법관 결선투표]\n선출 정원 경계 동률로 24시간 결선을 진행합니다. /후보 · /투표\n확정 당선자 " + e.Winners.Count + "명"); return; }
                else break;
                if (seats == 0) break;
            }
            e.State = "closed";
            e.NextReminderAt = 0;
            room.Notices.Add("[대법관 선거 종료]\n당선 " + e.Winners.Count + "명\n" + string.Join("\n", e.Candidates.Where(c => e.Winners.Contains(c.Number)).Select(c => c.Name)) + (runoff ? "\n결선에서도 동률인 잔여 자리는 공석으로 둡니다." : ""));
        }
        internal bool IsJudge(long chat, string user, long now)
        {
            JudgeRoom room; GovernanceRoster roster;
            if (!state.Rooms.TryGetValue(chat, out room) || !room.Enabled || !live.TryGetValue(chat, out roster) || now < roster.At || now - roster.At > 30 || !roster.Present.Contains(user)) return false;
            return room.Elections.Any(e => now >= e.End && now < e.TermEnd && e.Winners.Any(n => e.Candidates.Any(c => c.Number == n && c.User == user)) && !room.Revoked.Contains(e.Start + ":" + user));
        }
        internal void Leave(long chat, string user)
        {
            GovernanceRoster roster; if (live.TryGetValue(chat, out roster)) roster.Present.Remove(user);
            JudgeRoom room;
            if (!state.Rooms.TryGetValue(chat, out room) || !room.Enabled) return;
            var next = Copy(); bool changed = false;
            foreach (var e in next.Rooms[chat].Elections)
                if (e.Winners.Any(n => e.Candidates.Any(c => c.Number == n && c.User == user))) changed |= next.Rooms[chat].Revoked.Add(e.Start + ":" + user);
            if (changed) Commit(next);
        }
        internal string Handle(long chat, string user, string name, string verb, string argument, long now)
        {
            JudgeRoom room;
            if (verb == "/선거기록") return History(chat, argument);
            if (!state.Rooms.TryGetValue(chat, out room) || !room.Enabled) return "이 방은 대법관 선거가 활성화되어 있지 않습니다.";
            var election = room.Elections.LastOrDefault();
            if (verb == "/선거") return "[대법관 선거]\n" + (election == null ? "개최 기록 없음" : (election.Kind == "by" ? "보궐선거" : "정기선거") + " · 선출 " + election.Seats + "명\n상태: " + election.State + " · 선거인 " + election.Electors.Count + "명 · 최소 " + election.Threshold + "표\n마감: " + Local(election.End + (election.State == "runoff" ? 86400 : 0)).ToString("yyyy-MM-dd HH:mm")) + "\n다음 정기선거: " + Local(room.NextElection).ToString("yyyy-MM-dd HH:mm") + "\n1인 최대 " + (election == null ? 10 : election.State == "runoff" ? election.Seats - election.Winners.Count : election.Seats) + "명 승인투표 · 채팅에 선택 내용이 공개됩니다.";
            if (verb == "/대법관") return "[대법관]\n" + string.Join("\n", room.Elections.SelectMany(term => term.Candidates.Where(c => term.Winners.Contains(c.Number) && IsJudge(chat, c.User, now)).Select(c => c.Name)).Distinct());
            if (election == null) return "아직 선거가 없습니다.";
            if (verb == "/선거결과") return ElectionResult(election, room.Elections.Count);
            if (verb == "/후보") return "[후보]\n" + string.Join("\n", election.Candidates.Where(c => !c.Withdrawn && (election.State != "runoff" || election.Runoff.Contains(c.Number))).Select(c => c.Number + ". " + c.Name + "\n" + c.Pledge));
            bool runoff = election.State == "runoff";
            if (now < election.Start || now >= election.End + (runoff ? 86400 : 0) || election.State == "closed") return "현재 투표 기간이 아닙니다.";
            GovernanceRoster roster;
            if (!election.Electors.Contains(user) || !live.TryGetValue(chat, out roster) || now < roster.At || now - roster.At > 30 || !roster.Present.Contains(user)) return "이번 선거의 선거권이 없거나 현재 참여자 정보를 확인 중입니다.";
            var next = Copy(); var e = next.Rooms[chat].Elections.Last();
            var ballots = runoff ? e.RunoffBallots : e.Ballots;
            if (verb == "/투표취소") ballots.Remove(user);
            else if (verb == "/투표")
            {
                var parts = argument.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries); var numbers = new List<int>();
                foreach (string part in parts) { int number; if (!int.TryParse(part, out number)) return "형식: /투표 1 3 5"; numbers.Add(number); }
                if (numbers.Count < 1 || numbers.Count > (runoff ? e.Seats - e.Winners.Count : e.Seats) || numbers.Distinct().Count() != numbers.Count || numbers.Any(n => !e.Candidates.Any(c => c.Number == n && !c.Withdrawn) || runoff && !e.Runoff.Contains(n))) return "후보 번호 또는 선택 인원을 확인하세요. 같은 후보를 중복 선택할 수 없습니다.";
                ballots[user] = numbers;
            }
            else if (verb == "/출마")
            {
                if (e.Kind == "by" && Mandates(next.Rooms[chat], now).Contains(user)) return "재임 중인 대법관은 보궐선거에 출마할 수 없습니다.";
                if (runoff || argument.Length < 1 || argument.Length > 300 || argument.Any(char.IsControl)) return "본선 기간에 /출마 공약(1~300자) 형식으로 입력하세요.";
                var candidate = e.Candidates.FirstOrDefault(c => c.User == user);
                if (candidate == null) e.Candidates.Add(new ElectionCandidate { Number = e.Candidates.Count + 1, User = user, Name = name, Pledge = argument });
                else { if (candidate.Withdrawn) return "사퇴한 선거에는 다시 출마할 수 없습니다."; candidate.Pledge = argument; candidate.Name = name; }
            }
            else if (verb == "/출마취소")
            {
                var candidate = e.Candidates.FirstOrDefault(c => c.User == user);
                if (candidate == null) return "등록된 후보가 아닙니다.";
                candidate.Withdrawn = true;
            }
            else return "지원하지 않는 선거 명령입니다.";
            Commit(next); return "선거 기록을 저장했습니다.";
        }
        private static string ElectionResult(JudgeElection election, int number)
        {
            string phase = election.State == "closed" ? "종료" : election.State == "runoff" ? "결선 진행 중" : "투표 진행 중";
            return "[제" + number + "회 " + (election.Kind == "by" ? "보궐선거" : "정기선거") + "]\n" +
                Local(election.Start).ToString("yyyy-MM-dd HH:mm") + " ~ " + Local(election.End + (election.Runoff.Count > 0 ? 86400 : 0)).ToString("yyyy-MM-dd HH:mm") +
                "\n" + phase + " · 선출 " + election.Seats + "명 · 당선 " + election.Winners.Count + "명\n" +
                (election.Candidates.Count == 0 ? "등록 후보 없음" : string.Join("\n", election.Candidates.OrderBy(c => c.Number).Select(c =>
                    c.Number + ". " + c.Name + " · 본선 " + election.Ballots.Values.Count(v => v.Contains(c.Number)) + "표" +
                    (election.Runoff.Contains(c.Number) ? " · 결선 " + election.RunoffBallots.Values.Count(v => v.Contains(c.Number)) + "표" : "") +
                    " · " + (election.Winners.Contains(c.Number) ? "당선" : c.Withdrawn ? "사퇴" : election.State == "closed" ? "미당선" : "집계 중"))));
        }
        internal string History(long chat, string argument)
        {
            JudgeRoom room;
            if (!state.Rooms.TryGetValue(chat, out room) || room.Elections.Count == 0) return "아직 선거 기록이 없습니다.";
            if (!string.IsNullOrWhiteSpace(argument))
            {
                int number;
                if (!int.TryParse(argument, NumberStyles.None, CultureInfo.InvariantCulture, out number) || number < 1 || number > room.Elections.Count)
                    return "형식: /선거기록 또는 /선거기록 회차번호 (1~" + room.Elections.Count + ")";
                return ElectionResult(room.Elections[number - 1], number);
            }
            return "[선거기록 · 전체 " + room.Elections.Count + "회]\n" + string.Join("\n\n", room.Elections.Select((e, i) => ElectionResult(e, i + 1)));
        }
        internal static IEnumerable<string> Messages(string text)
        {
            if (text.Length <= 3000) { yield return text; yield break; }
            var chunks = new List<string>(); var part = new StringBuilder();
            foreach (string line in text.Split('\n'))
            {
                string rest = line + "\n";
                if (part.Length > 0 && part.Length + rest.Length > 3000) { chunks.Add(part.ToString().TrimEnd()); part.Clear(); }
                while (rest.Length > 3000)
                {
                    int length = char.IsHighSurrogate(rest[2999]) ? 2999 : 3000;
                    chunks.Add(rest.Substring(0, length)); rest = rest.Substring(length);
                }
                part.Append(rest);
            }
            if (part.Length > 0) chunks.Add(part.ToString().TrimEnd());
            for (int i = 0; i < chunks.Count; i++) yield return "[선거 안내 " + (i + 1) + "/" + chunks.Count + "]\n" + chunks[i];
        }
        internal bool QueueModeration(long chat, string key, string kind, string actor, string target, string reason, long now)
        {
            JudgeRoom current;
            if (state.Rooms.TryGetValue(chat, out current) && current.Moderation.Any(r => r.Key == key)) return false;
            var next = Copy();
            if (!next.Rooms.ContainsKey(chat)) next.Rooms.Add(chat, new JudgeRoom { ChatId = chat });
            next.Rooms[chat].Moderation.Add(new ModerationRecord { Key = key, Kind = kind, Actor = actor, Target = target, Reason = reason, At = now, Status = "queued" });
            Commit(next); return true;
        }
        internal void FinishModeration(long chat, string key, string error)
        {
            var next = Copy(); var row = next.Rooms[chat].Moderation.Single(r => r.Key == key);
            row.Status = error == null ? "macro_completed" : "macro_error"; row.Error = error;
            if (error != null) next.Rooms[chat].Notices.Add("관리 매크로 중단: " + error);
            Commit(next);
        }
        internal void RecordBans(long chat, IEnumerable<BanRecord> rows)
        {
            var next = Copy(); var room = next.Rooms[chat];
            foreach (var old in room.Bans) old.Listed = false;
            foreach (var row in rows)
            {
                var old = room.Bans.FirstOrDefault(b => b.User == row.User && b.KickedAt == row.KickedAt);
                if (old == null) { row.Number = room.Bans.Count + 1; room.Bans.Add(row); old = row; }
                old.Listed = true; old.Name = row.Name;
            }
            var listed = room.Bans.Where(b => b.Listed).OrderBy(b => b.Number).ToArray();
            room.Notices.Add("[내보낸 목록]\n" + (listed.Length == 0 ? "내보낸 상대가 없습니다." :
                string.Join("\n", listed.Select(b => b.Number + ". " + b.Name + " · " + Local(b.KickedAt).ToString("yyyy-MM-dd")))));
            Commit(next);
        }
        internal BanRecord FindBan(long chat, int number)
        { JudgeRoom room; return state.Rooms.TryGetValue(chat, out room) ? room.Bans.FirstOrDefault(b => b.Number == number && b.Listed) : null; }
        internal string PeekNotice(long chat) { JudgeRoom room; return state.Rooms.TryGetValue(chat, out room) ? room.Notices.FirstOrDefault() : null; }
        internal void TakeNotice(long chat) { var next = Copy(); next.Rooms[chat].Notices.RemoveAt(0); Commit(next); }
    }
}
