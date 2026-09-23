using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    // 각 행은 종류·고유 키·JSON 원문으로 저장합니다. JSON은 문자열 셀이므로 64비트 ID가 반올림되지 않습니다.
    internal sealed class OperationRecord
    {
        public string Kind, Key, Month, Text, Name, Source;
        public long ChatId, UserId, AuthorId, LogId, At, Count;
        public bool Read;
    }
    internal sealed class OperationsStore
    {
        internal static readonly string[] Headers = { "kind", "record_key", "payload_json" };
        internal readonly Dictionary<string, OperationRecord> Rows = new Dictionary<string, OperationRecord>();
        internal bool Dirty, ResetPending;
        internal static string Month(long at) { return DateTimeOffset.FromUnixTimeSeconds(at).ToOffset(TimeSpan.FromHours(9)).ToString("yyyy-MM", CultureInfo.InvariantCulture); }
        internal void Load(List<List<string>> rows)
        {
            foreach (var cells in rows)
            {
                if (cells.Count != 3) throw new FormatException("DB_Operations 행 형식 오류");
                var r = JsonConvert.DeserializeObject<OperationRecord>(cells[2]);
                if (r == null || r.Key != cells[1] || r.Kind != cells[0] || string.IsNullOrEmpty(r.Key) ||
                    !new[] { "month", "monthly", "popularity", "memo", "reentry", "point_reset" }.Contains(r.Kind))
                    throw new FormatException("DB_Operations 키 또는 종류 오류");
                DateTime month;
                if ((r.Kind == "month" || r.Kind == "monthly" || r.Kind == "popularity" || r.Kind == "point_reset") &&
                    !DateTime.TryParseExact(r.Month, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out month))
                    throw new FormatException("DB_Operations 월 형식 오류");
                if (r.Kind == "month" && r.Key != "current_month" || r.Kind != "month" && r.UserId <= 0 ||
                    r.At < 0 || (r.Kind == "monthly" || r.Kind == "popularity" || r.Kind == "memo" || r.Kind == "reentry") && r.ChatId <= 0 ||
                    r.Kind == "monthly" && r.Count < 0 || r.Kind == "memo" && (r.AuthorId <= 0 && r.Source != "local_operator" || string.IsNullOrWhiteSpace(r.Text)))
                    throw new FormatException("DB_Operations ID 또는 값 오류");
                Rows.Add(r.Key, r);
            }
            if (Rows.Count > 0 && !Rows.ContainsKey("current_month")) throw new FormatException("월 초기화 기준 행이 없습니다. DB_Operations를 확인하세요.");
            Dirty = ResetPending = false;
        }
        internal List<List<object>> ToRows()
        { return Rows.Values.OrderBy(r => r.Key, StringComparer.Ordinal).Select(r => new List<object> { r.Kind, r.Key, JsonConvert.SerializeObject(r) }).ToList(); }
        internal void Put(OperationRecord row) { Rows[row.Key] = row; Dirty = true; }

        // 최초 도입은 현재 월을 기준으로 등록합니다. 이후 월 전환 때만 전 사용자 포인트를 초기화합니다.
        internal bool AdvanceMonth(IEnumerable<User> users, long now)
        {
            string month = Month(now);
            OperationRecord state;
            if (!Rows.TryGetValue("current_month", out state))
            { Put(new OperationRecord { Kind = "month", Key = "current_month", Month = month, At = now }); ResetPending = true; return true; }
            if (string.CompareOrdinal(month, state.Month) <= 0) return ResetPending;
            foreach (var user in users)
            {
                Put(new OperationRecord { Kind = "point_reset", Key = "reset:" + month + ":" + user.UserId,
                    Month = month, UserId = user.UserId, Name = user.Nickname, Count = user.Point, At = now });
                user.Point = 0;
            }
            state.Month = month; state.At = now; Dirty = ResetPending = true; return true;
        }
        internal void CountMessage(ChatMessage message)
        {
            if (message.SendAt <= 0 || message.LogId <= 0 || message.AuthorId <= 0 || message.IsDirect || message.RoomEvent != ChatRoomEvent.None ||
                string.IsNullOrWhiteSpace(message.Message) || message.Message.TrimStart().StartsWith("/")) return;
            string month = Month(message.SendAt), key = "monthly:" + month + ":" + message.ChatId + ":" + message.AuthorId;
            OperationRecord row;
            if (!Rows.TryGetValue(key, out row)) row = new OperationRecord { Kind = "monthly", Key = key, Month = month, ChatId = message.ChatId, UserId = message.AuthorId };
            if (message.LogId <= row.LogId) return;
            row.Count = checked(row.Count + 1); row.LogId = message.LogId; row.Name = message.Nickname; row.At = message.SendAt; Put(row);
        }
        internal bool Reentry(ChatMessage message, IEnumerable<RoomEventRecord> history)
        {
            string key = "reentry:" + message.ChatId + ":" + message.LogId + ":" + message.AuthorId;
            if (Rows.ContainsKey(key)) return false;
            var previous = history.Where(e => e.ChatId == message.ChatId && e.UserId == message.AuthorId && e.LogId < message.LogId &&
                (e.Kind == "leave" || e.Kind == "kick")).OrderByDescending(e => e.LogId).FirstOrDefault();
            if (previous == null) return false;
            Put(new OperationRecord { Kind = "reentry", Key = key, ChatId = message.ChatId, UserId = message.AuthorId, LogId = message.LogId,
                Name = message.Nickname, At = message.SendAt, Text = "이전 이름: " + previous.Nickname + " / 마지막 " + (previous.Kind == "kick" ? "강퇴" : "퇴장") + ": " + LocalTime(previous.OccurredAt) });
            return true;
        }
        internal static string LocalTime(long at) { return DateTimeOffset.FromUnixTimeSeconds(at).ToOffset(TimeSpan.FromHours(9)).ToString("yyyy-MM-dd HH:mm:ss"); }
        internal bool AddPopularity(long chat, long user, long log, long amount, string name, long now)
        {
            string month = Month(now), key = "popularity:" + month + ":" + chat + ":" + user;
            OperationRecord row;
            if (!Rows.TryGetValue(key, out row)) row = new OperationRecord { Kind = "popularity", Key = key, ChatId = chat, UserId = user, Month = month };
            if (log <= row.LogId) return false;
            row.Count = checked(row.Count + amount); row.LogId = log; row.At = now; row.Name = name; Put(row); return true;
        }
        internal string Ranking(long chat, string month, bool popularity = false)
        {
            var rows = Rows.Values.Where(r => r.Kind == (popularity ? "popularity" : "monthly") && r.ChatId == chat && r.Month == month).OrderByDescending(r => r.Count).ThenBy(r => r.UserId).ToArray();
            return "[" + month + (popularity ? " 월간 인기도" : " 월간 채팅") + " 랭킹]\n" + (rows.Length == 0 ? "수집된 기록이 없습니다." : string.Join("\n", rows.Take(20).Select(r =>
                (1 + rows.Count(x => x.Count > r.Count)) + "위 " + r.Name + " · " + r.Count + (popularity ? "점" : "회")))) +
                (popularity ? "\n※ 이번 달 좋아요−싫어요 순증감 · 동점 공동 순위" : "\n※ 수집된 일반 채팅 기준 · 동점 공동 순위");
        }
    }
}
