using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Data.SQLite;

namespace KakaotalkBot
{
    public class Database
    {
        public const string UserSheetName = "DB_UserId";
        public bool UserStorageReady { get; private set; }
        public string UserStorageError { get; private set; }
        private static Database instance;
        public static Database Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Database();
                }

                return instance;
            }
        }

        private GoogleSheetHelper keywordSheet;
        private List<string> keywords = new List<string>();
        private List<List<string>> commands = new List<List<string>>();
        private List<User> userTable = new List<User>();
        private UserActivityStore activity = new UserActivityStore();
        internal RoomOperatorStore Operators = new RoomOperatorStore();
        internal OperationsStore Operations = new OperationsStore();
        public bool HasPendingActivity { get { return activity.Dirty || Operators.Dirty || Operations.Dirty; } }
        internal void ConfirmActivitySaved() { activity.MarkSaved(); Operators.Dirty = false; Operations.Dirty = false; Operations.ResetPending = false; }
        public bool IsRoomOperator(long chatId, long userId)
        { string error; return Operators.Check(chatId, userId, DateTimeOffset.UtcNow.ToUnixTimeSeconds(), out error); }
        public IReadOnlyList<RoomEventRecord> RoomEvents { get { return activity.Events.AsReadOnly(); } }
        public IReadOnlyList<NicknameRecord> NicknameHistory { get { return activity.Nicknames.AsReadOnly(); } }
        // 재실이 확인된 같은 방 사용자만 원본 멘션 후보로 사용합니다.
        internal string[] MentionFallbackNames(long chatId, long targetId)
        {
            return activity.Rooms.Values.Where(r => r.ChatId == chatId && r.UserId != targetId && r.IsPresent == true &&
                !string.IsNullOrWhiteSpace(r.Nickname) && "오픈채팅봇".IndexOf(r.Nickname, StringComparison.OrdinalIgnoreCase) < 0 &&
                !string.Equals(r.Nickname, "all", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(r => r.LastSeenAt).Select(r => r.Nickname).Distinct().Take(5).ToArray();
        }

        public RoomUser FindRoomUser(long chatId, long userId)
        { RoomUser value; return activity.Rooms.TryGetValue(Tuple.Create(chatId, userId), out value) ? value : null; }
        private readonly HashSet<Tuple<long, long, long>> departureEvents = new HashSet<Tuple<long, long, long>>();
        private readonly Queue<Tuple<long, long, long>> departureEventOrder = new Queue<Tuple<long, long, long>>();
        private List<Quiz> commonSenses = new List<Quiz>();
        private List<Topic> topics = new List<Topic>();

        private RandomNumberGenerator random;
        private byte[] randomBytes = new byte[4];
        private int currentAnswerIndex = -1;
        private string[] categoryOrder = new string[]
        {
            "🗳정치", "🎞역사", "📊경제", "👨‍🏫인문학", "🌍기타"
        };

        public List<List<string>> Commands { get { return commands; } }
        public List<string> Keywords { get { return keywords; } }
        public IReadOnlyList<User> UserTable { get { return userTable.AsReadOnly(); } }
        public List<Quiz> CommonSenses { get { return commonSenses; } }
        public int CurrentAnswerIndex { get {  return currentAnswerIndex; }  set { currentAnswerIndex = value; } }
        public List<Topic> Topics { get { return topics; } }
        public string[] CategoryOrder { get { return categoryOrder; } }


        private Database()
        {

        }

        public void Initialize(string applicationName, string spreadsheetId)
        {
            UserStorageReady = false;
            random = RandomNumberGenerator.Create();
            keywordSheet = new GoogleSheetHelper(applicationName, spreadsheetId);
            commands = GetCommanads();
            keywords = GetKeywords();
            try
            {
                keywordSheet.EnsureActivitySchema();
                userTable = GetUserTable();
                var next = new UserActivityStore();
                next.Load(keywordSheet.ReadActivityTable("DB_RoomUsers", UserActivityStore.RoomHeaders),
                    keywordSheet.ReadActivityTable("DB_RoomEvents", UserActivityStore.EventHeaders),
                    keywordSheet.ReadActivityTable("DB_NicknameHistory", UserActivityStore.NicknameHeaders));
                var ids = new HashSet<long>(userTable.Select(u => u.UserId));
                var nextOperators = new RoomOperatorStore();
                nextOperators.Load(keywordSheet.ReadActivityTable("DB_RoomOperators", RoomOperatorStore.Headers));
                var nextOperations = new OperationsStore();
                nextOperations.Load(keywordSheet.ReadActivityTable("DB_Operations", OperationsStore.Headers));
                if (next.Rooms.Values.Any(r => !ids.Contains(r.UserId)) || next.Events.Any(r => !ids.Contains(r.UserId)) || next.Nicknames.Any(r => !ids.Contains(r.UserId)))
                    throw new FormatException("활동 DB에 사용자 DB에 없는 user_id가 있습니다.");
                activity = next; Operators = nextOperators; UserStorageReady = true; UserStorageError = null;
                Operations = nextOperations;
                MaintainMonth();
            }
            catch (Exception error) { UserStorageReady = false; UserStorageError = "사용자 DB 읽기 실패: " + error.Message; throw; }
            commonSenses = GetCommonSenses();
            topics = GetTopic();
        }

        public void UpdateCommands()
        {
            commands = GetCommanads();
            keywords = GetKeywords();
        }

        public void ResetAttendance()
        {
            List<List<object>> users = new List<List<object>>();
            foreach (User user in userTable)
            {
                user.TakeAttendance = false;
                users.Add(user.ToRow());
            }
        }

        public void UpdateUserTable()
        {
            if (!UserStorageReady) throw new InvalidOperationException("사용자 DB를 정상적으로 읽기 전에는 저장할 수 없습니다.");
            try
            {
                keywordSheet.WriteUserData(userTable.Select(user => user.ToRow()).ToList(), activity, Operators, Operations);
                UserStorageError = null;
            }
            catch (Exception error) { UserStorageError = "사용자 DB 저장 실패: " + error.Message; throw; }
        }

        public void UpdateCommonSenses()
        {
            var list = GetCommonSenses();
            if (list != null && list.Count != 0)
            {
                commonSenses = list;

            }
        }

        public void UpdateTopic()
        {
            topics = GetTopic();
        }

        public User AddUser(long userId, string nickname)
        {
            var user = GetOrAddUser(userId, nickname);
            // 전역 닉네임은 이전 데이터 보존용입니다. 방별 표시 이름은 DB_RoomUsers에서 조회합니다.
            return user;
        }

        // 멘션으로 처음 확인한 사용자도 등록한다. 기존 표시 이름과 자산은 유지한다.
        public User GetOrAddUser(long userId, string nickname = null)
        {
            if (userId <= 0) throw new ArgumentOutOfRangeException("userId");
            User user;
            if (!FindUser(userId, out user))
            {
                user = new User(userId) { Nickname = string.IsNullOrWhiteSpace(nickname) ? "사용자 " + userId : nickname };
                userTable.Add(user);
            }
            else if (string.IsNullOrWhiteSpace(user.Nickname) && !string.IsNullOrWhiteSpace(nickname)) user.Nickname = nickname;
            return user;
        }

        public bool FindUser(long userId)
        {
            User user;
            return FindUser(userId, out user);
        }

        public bool FindUser(long userId, out User user)
        {
            user = null;
            if (userId <= 0) return false;
            foreach (User u in userTable)
            {
                if (u.UserId == userId)
                {
                    user = u;
                    return true;
                }
            }
            user = null;
            return false;
        }

        public bool CheckAttendance(long userId, string nickname)
        {
            var user = AddUser(userId, nickname);
            // 봇 재시작이나 출석 표시 초기화 뒤에도 같은 날짜의 보상은 다시 지급하지 않는다.
            if (user.AttendanceAt.Date == DateTime.Today) { user.TakeAttendance = true; return true; }
            int points = checked(user.Point + 10);
            user.TakeAttendance = true;
            user.AttendanceAt = DateTime.Now;
            user.Point = points;
            return false;
        }

        public bool ChangePopularity(long authorId, long targetId, int amount, bool increase, out string error, long chatId = 0, long logId = 0)
        {
            error = null;
            if (authorId <= 0 || targetId <= 0 || amount <= 0) error = "올바른 사용자 ID와 양의 포인트가 필요합니다.";
            else if (authorId == targetId) error = "자신에게 할 수 없는 명령입니다.";
            else
            {
                var author = GetOrAddUser(authorId);
                var target = GetOrAddUser(targetId);
                if (author.Point < amount) { error = "포인트가 부족합니다.\n남은 포인트: " + author.Point; return false; }
                long popularity = (long)target.Popularity + (increase ? (long)amount : -(long)amount);
                if (popularity < int.MinValue || popularity > int.MaxValue) error = "인기도의 저장 범위를 초과합니다.";
                else
                {
                    if (chatId > 0 && !Operations.AddPopularity(chatId, targetId, logId, increase ? amount : -amount, ChatUserName(chatId, targetId), DateTimeOffset.UtcNow.ToUnixTimeSeconds()))
                    { error = "이미 처리한 인기도 명령입니다."; return false; }
                    author.Point -= amount; target.Popularity = (int)popularity;
                    if (chatId > 0) UpdateUserTable();
                    return true;
                }
            }
            return false;
        }

        public List<User> GetPopularityRank()
        {
            return userTable.OrderByDescending(x => x.Popularity).ToList();
        }

        // 시스템 이벤트의 대상 ID를 집계한다. 같은 방·로그·사용자는 한 번만 반영한다.
        public bool RecordDeparture(long userId, string nickname, long chatId, long logId, bool kicked, long occurredAt = 0)
        {
            if (userId <= 0 || chatId <= 0 || logId <= 0) throw new ArgumentOutOfRangeException("userId", "사용자·채팅방·로그 ID가 필요합니다.");
            var key = Tuple.Create(chatId, logId, userId);
            if (departureEvents.Contains(key) || activity.HasEvent(chatId, userId, logId)) return false;
            var user = AddUser(userId, nickname);
            if (kicked) user.KickCount = checked(user.KickCount + 1);
            else user.LeaveCount = checked(user.LeaveCount + 1);
            activity.RecordEvent(chatId, userId, logId, kicked ? "kick" : "leave", nickname, occurredAt);
            departureEvents.Add(key);
            departureEventOrder.Enqueue(key);
            // 더 오래된 로그의 재처리는 수신 커서가 차단한다.
            if (departureEventOrder.Count > 4096) departureEvents.Remove(departureEventOrder.Dequeue());
            return true;
        }

        internal bool ObserveMessage(ChatMessage message)
        {
            var user = GetOrAddUser(message.AuthorId, message.Nickname);
            if (message.RoomEvent == ChatRoomEvent.Join)
            {
                Operations.Reentry(message, activity.Events);
                activity.RecordEvent(message.ChatId, message.AuthorId, message.LogId, "join", message.Nickname, message.SendAt);
                return true;
            }
            if (message.RoomEvent != ChatRoomEvent.None) return false;
            if (!message.IsOwn) Operations.CountMessage(message);
            bool renamed = activity.Observe(message.ChatId, message.AuthorId, message.Nickname, true,
                message.SendAt, message.LogId, "chat");
            // 일반 채팅과 답글에만 경험치를 지급합니다. 명령어·시스템 이벤트는 제외합니다.
            if (!message.IsOwn && !string.IsNullOrWhiteSpace(message.Message) && !message.Message.TrimStart().StartsWith("/"))
                activity.AwardExperience(user, message.ChatId, message.LogId, message.SendAt);
            return renamed;
        }

        internal void ObserveRoster(ChatRosterSnapshot roster)
        {
            Operators.Observe(roster);
            foreach (var member in roster.Members)
            {
                var user = GetOrAddUser(member.UserId, member.Nickname);
                activity.Observe(roster.ChatId, member.UserId, member.Nickname, member.IsPresent,
                    roster.ObservedAt, roster.LogId, "profile");

            }
        }

        public string DescribeActivity(long chatId, long userId)
        {
            var room = FindRoomUser(chatId, userId);
            string present = room == null || !room.IsPresent.HasValue ? "미확인" : room.IsPresent.Value ? "참여 중" : "퇴장 상태";
            var events = RoomEvents.Where(e => e.ChatId == chatId && e.UserId == userId).OrderByDescending(e => e.OccurredAt).ThenByDescending(e => e.LogId).Take(5);
            var names = NicknameHistory.Where(e => e.ChatId == chatId && e.UserId == userId).OrderByDescending(e => e.ObservedAt).Take(5);
            Func<long, string> time = seconds => DateTimeOffset.FromUnixTimeSeconds(seconds).ToLocalTime().ToString("MM-dd HH:mm");
            return "현재 방: " + present + "\n[최근 입퇴장]\n" + string.Join("\n", events.Select(e => time(e.OccurredAt) + " " + (e.Kind == "join" ? "입장" : e.Kind == "kick" ? "강퇴" : "퇴장"))) +
                "\n[최근 닉네임 변경]\n" + string.Join("\n", names.Select(e => time(e.ObservedAt) + " " + e.Before + " → " + e.After));
        }

        public void MaintainMonth()
        {
            if (!UserStorageReady) return;
            if (Operations.AdvanceMonth(userTable, DateTimeOffset.UtcNow.ToUnixTimeSeconds())) UpdateUserTable();
        }

        public void AddOperatorMemo(long chat, long author, long target, string text, long log = 0)
        {
            if (!UserStorageReady || !IsRoomOperator(chat, author)) throw new InvalidOperationException("현재 방의 운영진 권한을 확인할 수 없습니다. 수신 상태와 작성자 ID를 확인하세요.");
            SaveMemo(chat, author, target, text, log, "chat_operator");
        }
        // 로컬 운영 화면은 이 PC의 관리자가 사용합니다. 특정 카카오 사용자로 가장하지 않습니다.
        internal void AddLocalMemo(long chat, long target, string text)
        { SaveMemo(chat, 0, target, text, 0, "local_operator"); }
        private void SaveMemo(long chat, long author, long target, string text, long log, string source)
        {
            if (!UserStorageReady || chat <= 0) throw new InvalidOperationException("사용자 DB와 채팅방을 먼저 확인하세요.");
            if (target <= 0 || string.IsNullOrWhiteSpace(text) || text.Trim().Length > 1000) throw new ArgumentException("대상 ID와 1~1000자의 메모를 입력하세요.");
            GetOrAddUser(target);
            string key = log > 0 ? "memo:" + chat + ":" + log : "memo:" + Guid.NewGuid().ToString("N");
            if (!Operations.Rows.ContainsKey(key)) Operations.Put(new OperationRecord { Kind = "memo", Key = key, ChatId = chat, UserId = target,
                AuthorId = author, LogId = log, Text = text.Trim(), Source = source, At = DateTimeOffset.UtcNow.ToUnixTimeSeconds() });
            UpdateUserTable();
        }
        public string OperatorHistory(long chat, long target, bool includeUserIds = true)
        {
            User user; FindUser(target, out user);
            return ChatUserName(chat, target) + (includeUserIds ? " · ID " + target : "") +
                "\n[입퇴장 이력]\n" + string.Join("\n", activity.Events.Where(r => r.ChatId == chat && r.UserId == target).OrderByDescending(r => r.LogId)
                    .Select(r => OperationsStore.LocalTime(r.OccurredAt) + " " + (r.Kind == "join" ? "입장" : r.Kind == "kick" ? "강퇴" : "퇴장") + " · " + r.Nickname)) +
                "\n[닉네임 변경 이력]\n" + string.Join("\n", activity.Nicknames.Where(r => r.ChatId == chat && r.UserId == target).OrderByDescending(r => r.ObservedAt)
                    .Select(r => OperationsStore.LocalTime(r.ObservedAt) + " " + r.Before + " → " + r.After)) +
                "\n[운영 메모]\n" + string.Join("\n", Operations.Rows.Values.Where(r => r.Kind == "memo" && r.ChatId == chat && r.UserId == target)
                    .OrderByDescending(r => r.At).Select(r => OperationsStore.LocalTime(r.At) + " / 작성자 " + (r.Source == "local_operator" ? "로컬 운영자" : includeUserIds ? r.AuthorId.ToString() : ChatUserName(chat, r.AuthorId)) + "\n" + r.Text));
        }
        public string ChatUserName(long chatId, long userId, string observedNickname = null)
        {
            var room = FindRoomUser(chatId, userId);
            if (room != null && !string.IsNullOrWhiteSpace(room.Nickname)) return room.Nickname;
            // 현재 명령의 멘션처럼 해당 방에서 얻은 이름만 보조값으로 허용합니다.
            return !string.IsNullOrWhiteSpace(observedNickname) ? observedNickname : "이름 미확인 사용자";
        }

        // 이전 버전에서 만든 대기 알림과 ID 기반 임시 닉네임도 전송 직전에 정리합니다.
        internal string ForChat(string text)
        {
            if (text == null) return null;
            text = System.Text.RegularExpressions.Regex.Replace(text, @"(?m)^\s*(?:사용자 ID|요청자 ID|ID):\s*\d+\s*\r?$", "");
            text = System.Text.RegularExpressions.Regex.Replace(text, @" · ID \d+", "");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"사용자 \d+", "이름 미확인 사용자");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"(작성자:?\s*)\d+", "$1운영진");
            var ids = new HashSet<string>(userTable.Select(u => u.UserId.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            // 긴 ID가 메모 본문 등에 남아 있어도 채팅 메시지로 노출하지 않습니다.
            return System.Text.RegularExpressions.Regex.Replace(text, @"(?<!\d)\d{10,19}(?!\d)", match => ids.Contains(match.Value) ? "[사용자]" : match.Value);
        }

        private User[] LevelRankUsers(long chat)
        {
            return userTable.Where(u => { var room = FindRoomUser(chat, u.UserId); return room != null && room.IsPresent == true; })
                .OrderByDescending(u => u.Experience).ThenBy(u => u.UserId).ToArray();
        }

        public string LevelRanking(long chat)
        {
            var users = LevelRankUsers(chat);
            return "[레벨 랭킹]\n" + (users.Length == 0 ? "참여 중으로 확인된 사용자가 없습니다." :
                string.Join("\n", users.Take(20).Select(u => (1 + users.Count(other => other.Experience > u.Experience)) +
                    "위 " + ChatUserName(chat, u.UserId) + " · Lv." + u.Level + " · 경험치 " + u.Experience))) +
                "\n※ 현재 방 참여 확인 사용자 · 누적 경험치 기준 · 동점 공동 순위";
        }

        public string PersonalRankings(long chat, long userId, string month)
        {
            var rows = Operations.Rows.Values.Where(r => r.Kind == "monthly" && r.ChatId == chat && r.Month == month).ToArray();
            var mine = rows.FirstOrDefault(r => r.UserId == userId);
            string monthly = mine == null || mine.Count <= 0 ? "순위 없음 (0회)" :
                (1 + rows.Count(r => r.Count > mine.Count)) + "위 / " + rows.Count(r => r.Count > 0) + "명 (" + mine.Count + "회)";
            var users = LevelRankUsers(chat);
            var user = users.FirstOrDefault(u => u.UserId == userId);
            string level = user == null ? "순위 없음 (현재 방 참여 미확인)" :
                (1 + users.Count(u => u.Experience > user.Experience)) + "위 / " + users.Length + "명";
            return "이달 채팅 순위 (" + month + "): " + monthly + "\n레벨 순위: " + level;
        }

        public string RoomStatistics(long chat, string month)
        {
            var messages = Operations.Rows.Values.Where(r => r.Kind == "monthly" && r.ChatId == chat && r.Month == month).ToArray();
            var events = activity.Events.Where(e => e.ChatId == chat && OperationsStore.Month(e.OccurredAt) == month).ToArray();
            return "[" + month + " 방 통계]\n일반 채팅: " + messages.Sum(r => r.Count) + "회\n활동 사용자: " + messages.Length +
                "명\n입장: " + events.Count(e => e.Kind == "join") + "회 / 퇴장: " + events.Count(e => e.Kind == "leave") +
                "회 / 강퇴: " + events.Count(e => e.Kind == "kick") + "회\n현재 참여 확인: " + activity.Rooms.Values.Count(r => r.ChatId == chat && r.IsPresent == true) +
                "명\n※ 수신한 기록과 마지막 참여자 목록 기준";
        }

        private List<User> GetUserTable()
        {
            var db = keywordSheet.ReadUserTable();
            List<User> users = new List<User>();
            var ids = new HashSet<long>();
            foreach (var row in db)
            {
                var user = User.ToUser(row);
                if (!ids.Add(user.UserId)) throw new FormatException("사용자 DB에 중복된 user_id가 있습니다: " + user.UserId);
                users.Add(user);
            }
            return users;
        }

        private List<Topic> GetTopic()
        {
            var db = keywordSheet.ReadAllFromSheet("Topic");
            List<Topic> users = new List<Topic>();
            for (int i = 0; i < db.Count; i++)
            {
                if (i == 0) continue;
                var row = db[i];
                users.Add(Topic.ToTopic(row));
            }
            return users;
        }

        private List<Quiz> GetCommonSenses()
        {
            var db = keywordSheet.ReadAllFromSheet("상식퀴즈");
            List<Quiz> users = new List<Quiz>();
            for (int i = 0; i < db.Count; i++) 
            {
                if (i == 0) continue;
                var row = db[i];
                users.Add(Quiz.ToCommonSense(row));
            }
            return users;
        }

        private List<List<string>> GetCommanads()
        {
            return keywordSheet.ReadCommandTables();
        }

        public List<string> GetKeywords()
        {
            List<string> keywordList = new List<string>();
            foreach (var row in commands)
            {
                keywordList.Add(row[0]);
            }

            return keywordList;
        }

        public string GetAnswer(string keyword)
        {
            var keywords = commands;
            foreach (var row in keywords)
            {
                if (row[0] == keyword && row.Count > 1) 
                {
                    return (keyword == "/?" || keyword == "/명령어") && !row[1].Contains("/레벨랭킹")
                        ? row[1].Replace("/채팅랭킹", "/채팅랭킹\n/레벨랭킹") : row[1];
                }
            }

            return null;
        }

        public string[] GetAnswers(string keyword)
        {
            var keywords = commands;
            foreach (var row in keywords)
            {
                if (row[0] == keyword && row.Count > 1)
                {
                    return row[1].Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                }
            }

            return null;
        }

        public void SetNextCommonSense()
        {
            random.GetBytes(randomBytes);
            int randomValue = BitConverter.ToInt32(randomBytes, 0);
            randomValue = Math.Abs(randomValue);
            currentAnswerIndex = randomValue % commonSenses.Count;
        }

        public void ResetCommonSense()
        {
            currentAnswerIndex = -1;
        }

        public string GetCommonSenseText()
        {
            if(currentAnswerIndex < 0 )
            {
                return string.Empty;
            }
            Quiz cs = commonSenses[currentAnswerIndex];
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("[⁉️상식퀴즈⁉️]");
            sb.AppendLine($"📜{cs.Question}");
            sb.AppendLine($"♻️분류: {cs.Category}");
            sb.AppendLine($"🎢난이도: {cs.Difficulty}");
            sb.Append($"💡힌트: {cs.Hint}");
            
            return sb.ToString();
        }

        public Quiz GetCurrentQuiz()
        {
            if (currentAnswerIndex < 0) return null;

            Quiz cs = commonSenses[currentAnswerIndex];
            return cs;
        }

        public List<Topic> GetOrderedTopics()
        {
            return topics.OrderByDescending(x => x.CreatedAt).OrderBy(x => Array.IndexOf(categoryOrder, x.Category)).ToList();
        }

        public int GetTotalContribution()
        {
            int totalContribution = 0;
            foreach(User user in userTable)
            {
                totalContribution += user.Contribution;
            }
            return totalContribution;
        }

        public List<User> GetContributionRank()
        {
            return userTable.OrderByDescending(x => x.Contribution).ToList();
        }
    }
}
