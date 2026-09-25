using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    internal sealed class QuestReward
    {
        public string Key;
        public int Points;
        public long At, ChatId, LogId;
    }

    internal sealed class DailyQuestRecord
    {
        public string Key, Day, LastTextHash;
        public long UserId, LastChatAt;
        public int ChatCount, QuizCount;
        public bool Attendance, Quiz; // Quiz는 기존 기록 호환용입니다.
        public Dictionary<long, long> QuizCursors = new Dictionary<long, long>();
        public Dictionary<long, long> ChatCursors = new Dictionary<long, long>();
        public Dictionary<string, QuestReward> Rewards = new Dictionary<string, QuestReward>();
    }

    // 날짜별 기록을 보존하므로 자정에 행을 지우거나 진행도를 덮어쓰지 않습니다.
    internal sealed class DailyQuestStore
    {
        internal static readonly string[] Headers = { "quest_key", "payload_json" };
        internal readonly Dictionary<string, DailyQuestRecord> Rows = new Dictionary<string, DailyQuestRecord>();
        internal bool Dirty;
        internal static string Day(long at)
        { return DateTimeOffset.FromUnixTimeSeconds(at).ToOffset(TimeSpan.FromHours(9)).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); }
        private static string Key(long user, string day) { return day + ":" + user.ToString(CultureInfo.InvariantCulture); }

        internal void Load(List<List<string>> rows)
        {
            foreach (var cells in rows)
            {
                var row = cells.Count == 2 ? JsonConvert.DeserializeObject<DailyQuestRecord>(cells[1]) : null;
                DateTime date;
                if (row == null || row.UserId <= 0 || !DateTime.TryParseExact(row.Day, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
                    row.Key != cells[0] || row.Key != Key(row.UserId, row.Day) || row.ChatCount < 0 || row.ChatCount > 20 ||
                    row.ChatCursors == null || row.QuizCursors == null || row.QuizCount < 0 || row.QuizCount > 3 || row.Rewards == null || row.LastChatAt < 0)
                    throw new FormatException("일일 퀘스트 기록 형식 오류");
                foreach (var pair in row.Rewards)
                {
                    int expected = pair.Key == "quiz" ? 20 : 10;
                    if (!new[] { "attendance", "chat", "quiz", "complete" }.Contains(pair.Key) || pair.Value == null ||
                        pair.Value.Key != row.Key + ":" + pair.Key || pair.Value.Points != expected || pair.Value.At <= 0 || Day(pair.Value.At) != row.Day)
                        throw new FormatException("퀘스트 보상 지급 기록 오류");
                }
                if (row.ChatCursors.Any(p => p.Key <= 0 || p.Value <= 0)) throw new FormatException("퀘스트 채팅 기준점 오류");
                // 기존 완료 여부만 있는 기록은 최소 한 번의 정답으로 이어갑니다. 이미 지급한 보상은 유지합니다.
                if (row.Quiz && row.QuizCount == 0) row.QuizCount = 1;
                QuestReward oldQuiz;
                if (row.Rewards.TryGetValue("quiz", out oldQuiz) && oldQuiz.ChatId > 0 && oldQuiz.LogId > 0 && !row.QuizCursors.ContainsKey(oldQuiz.ChatId))
                    row.QuizCursors[oldQuiz.ChatId] = oldQuiz.LogId;
                if (row.QuizCursors.Any(p => p.Key <= 0 || p.Value <= 0)) throw new FormatException("퀘스트 정답 기준점 오류");
                Rows.Add(row.Key, row);
            }
            Dirty = false;
        }

        internal List<List<object>> ToRows()
        { return Rows.Values.OrderBy(r => r.Key, StringComparer.Ordinal).Select(r => new List<object> { r.Key, JsonConvert.SerializeObject(r) }).ToList(); }

        private DailyQuestRecord Get(long user, long now)
        {
            if (user <= 0) throw new ArgumentException("사용자 ID가 필요합니다.");
            string day = Day(now), key = Key(user, day);
            DailyQuestRecord row;
            if (!Rows.TryGetValue(key, out row))
            { row = new DailyQuestRecord { Key = key, Day = day, UserId = user }; Rows.Add(key, row); Dirty = true; }
            return row;
        }

        internal string Complete(User user, string kind, long chat, long log, long now)
        {
            if (kind != "attendance" && kind != "quiz") throw new ArgumentException("알 수 없는 퀘스트");
            var row = Get(user.UserId, now);
            if (kind == "attendance") row.Attendance = true;
            else
            {
                long last;
                if (chat <= 0 || log <= 0) throw new ArgumentException("정답의 방과 메시지 ID가 필요합니다.");
                if (row.QuizCursors.TryGetValue(chat, out last) && log <= last) return "";
                row.QuizCursors[chat] = log;
                row.QuizCount = Math.Min(3, row.QuizCount + 1);
                row.Quiz = true;
            }
            Dirty = true;
            return Award(row, user, chat, log, now);
        }

        internal bool CountChat(User user, ChatMessage message, long now)
        {
            if (message.IsOwn || message.IsDirect || message.RoomEvent != ChatRoomEvent.None || message.EventCommand != null ||
                message.AuthorId != user.UserId || message.ChatId <= 0 || message.LogId <= 0 || message.SendAt <= 0 ||
                message.SendAt > now || Day(message.SendAt) != Day(now) || string.IsNullOrWhiteSpace(message.Message) || message.Message.TrimStart().StartsWith("/")) return false;
            var row = Get(user.UserId, now);
            long last;
            if (row.ChatCursors.TryGetValue(message.ChatId, out last) && message.LogId <= last) return false;
            if (row.ChatCount >= 10) return Award(row, user, message.ChatId, message.LogId, now).Length > 0;
            row.ChatCursors[message.ChatId] = message.LogId; Dirty = true;
            string hash;
            using (var sha = SHA256.Create()) hash = Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(message.Message.Trim())));
            if (row.LastChatAt != 0 && message.SendAt - row.LastChatAt < 60 || row.LastTextHash == hash) return false;
            row.ChatCount++; row.LastChatAt = message.SendAt; row.LastTextHash = hash;
            Award(row, user, message.ChatId, message.LogId, now);
            return true;
        }

        private string Award(DailyQuestRecord row, User user, long chat, long log, long now)
        {
            var available = new Dictionary<string, int>();
            if (row.Attendance) available.Add("attendance", 10);
            if (row.ChatCount >= 10) available.Add("chat", 10);
            if (row.QuizCount >= 3) available.Add("quiz", 20);
            if (row.Attendance && row.ChatCount >= 10 && (row.QuizCount >= 3 || row.Rewards.ContainsKey("quiz"))) available.Add("complete", 10);
            var pending = available.Where(p => !row.Rewards.ContainsKey(p.Key)).ToArray();
            int balance = checked(user.Point + pending.Sum(p => p.Value));
            foreach (var reward in pending)
                row.Rewards.Add(reward.Key, new QuestReward { Key = row.Key + ":" + reward.Key, Points = reward.Value, At = now, ChatId = chat, LogId = log });
            user.Point = balance;
            if (pending.Length > 0) Dirty = true;
            return string.Join("\n", pending.Select(p => "일일 " + (p.Key == "attendance" ? "출석" : p.Key == "quiz" ? "퀴즈" : p.Key == "chat" ? "대화 참여" : "완주") + " 퀘스트: +" + p.Value + "포인트"));
        }

        internal string Status(long user, long now)
        {
            DailyQuestRecord row;
            string day = Day(now);
            if (!Rows.TryGetValue(Key(user, day), out row)) row = new DailyQuestRecord();
            Func<string, string, string> line = (id, label) => (row.Rewards.ContainsKey(id) ? "✅ " : "⬜ ") + label +
                (row.Rewards.ContainsKey(id) ? " · " + row.Rewards[id].Points + "P 지급" : "");
            return "[오늘의 퀘스트 · " + day + "]\n" + line("attendance", "오늘도 출석 " + (row.Attendance ? "1" : "0") + "/1회 · 보상 10P") + "\n" +
                line("chat", "대화 참여 " + Math.Min(10, row.ChatCount) + "/10회 · 보상 10P") + "\n" + line("quiz", "오늘의 지식 " + row.QuizCount + "/3회 · 보상 20P") + "\n" +
                line("complete", "일일 완주 " + ((row.Attendance ? 1 : 0) + (row.QuizCount >= 3 || row.Rewards.ContainsKey("quiz") ? 1 : 0) + (row.ChatCount >= 10 ? 1 : 0)) + "/3개 · 보상 10P") +
                "\n\n오늘 받은 퀘스트 보상: " + row.Rewards.Values.Sum(r => r.Points) + "P / 50P\n초기화: 매일 00:00 (한국 시간)\n※ 일반 채팅은 60초에 1회, 직전 인정 내용과 같으면 제외 · 방 간 진행 공유";
        }
    }
}
