using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    internal sealed class RoomTitle
    {
        public string Key, Title, Source, Reason;
        public long ChatId, UserId, AuthorId, GrantedAt, ExpiresAt, RevokedAt, RevokedBy;
        public bool Active(long now) { return GrantedAt <= now && RevokedAt == 0 && (ExpiresAt == 0 || now < ExpiresAt); }
    }

    // 방과 사용자 ID로 칭호를 관리하며 해제·만료 후에도 원본 행을 보존합니다.
    internal sealed class RoomTitleStore
    {
        internal static readonly string[] Headers = { "title_key", "payload_json" };
        internal readonly Dictionary<string, RoomTitle> Rows = new Dictionary<string, RoomTitle>();
        internal bool Dirty;
        internal void Load(List<List<string>> rows)
        {
            foreach (var cells in rows)
            {
                if (cells.Count != 2) throw new FormatException("네임드 행 형식 오류");
                var row = JsonConvert.DeserializeObject<RoomTitle>(cells[1]);
                if (row == null || row.Key != cells[0] || string.IsNullOrWhiteSpace(row.Key) || row.ChatId <= 0 || row.UserId <= 0 ||
                    string.IsNullOrWhiteSpace(row.Title) || row.Title.Length > 30 || row.GrantedAt <= 0 || row.RevokedAt < 0 ||
                    row.ExpiresAt != 0 && row.ExpiresAt <= row.GrantedAt ||
                    row.Source != "manual" && row.Source != "monthly" || row.Source == "manual" && row.AuthorId <= 0)
                    throw new FormatException("네임드 값 또는 ID 오류");
                Rows.Add(row.Key, row);
            }
            Dirty = false;
        }
        internal List<List<object>> ToRows()
        { return Rows.Values.OrderBy(r => r.Key, StringComparer.Ordinal).Select(r => new List<object> { r.Key, JsonConvert.SerializeObject(r) }).ToList(); }
        internal IEnumerable<RoomTitle> Active(long chat, long user, long now)
        { return Rows.Values.Where(r => r.ChatId == chat && r.UserId == user && r.Active(now)); }
        internal bool Grant(long chat, long user, long author, long log, string title, int days, long now)
        {
            if (chat <= 0 || user <= 0 || author <= 0 || log <= 0 || now <= 0 || string.IsNullOrWhiteSpace(title) ||
                title.Length > 30 || title.Any(char.IsControl) || days < 0 || days > 3650)
                throw new ArgumentException("칭호는 1~30자, 기간은 1~3650일로 입력하세요.");
            string key = "manual:" + chat + ":" + log;
            if (Rows.ContainsKey(key) || Active(chat, user, now).Any(r => r.Title == title)) return false;
            Rows.Add(key, new RoomTitle { Key = key, ChatId = chat, UserId = user, AuthorId = author, Title = title,
                Source = "manual", Reason = "운영진 지정", GrantedAt = now, ExpiresAt = days == 0 ? 0 : checked(now + days * 86400L) });
            Dirty = true; return true;
        }
        internal bool Revoke(long chat, long user, string title, long now, long author)
        {
            var rows = Active(chat, user, now).Where(r => r.Title == title).ToArray();
            foreach (var row in rows) { row.RevokedAt = now; row.RevokedBy = author; }
            if (rows.Length > 0) Dirty = true;
            return rows.Length > 0;
        }
        internal void AwardMonthly(OperationsStore operations, long now)
        {
            string current = OperationsStore.Month(now);
            var first = DateTime.SpecifyKind(DateTime.ParseExact(current, "yyyy-MM", CultureInfo.InvariantCulture), DateTimeKind.Unspecified);
            long start = new DateTimeOffset(first, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            long end = new DateTimeOffset(first.AddMonths(1), TimeSpan.FromHours(9)).ToUnixTimeSeconds();
            string previous = first.AddMonths(-1).ToString("yyyy-MM", CultureInfo.InvariantCulture);
            // 동점 3위까지 모두 부여합니다. 양수 기록만 대상으로 하며 과거 집계는 보존합니다.
            foreach (var group in operations.Rows.Values.Where(r => (r.Kind == "monthly" || r.Kind == "popularity") && r.Month == previous && r.Count > 0)
                .GroupBy(r => new { r.ChatId, r.Kind }))
            {
                var candidates = group.ToArray();
                foreach (var winner in candidates.Where(r => candidates.Count(other => other.Count > r.Count) < 3))
                {
                    string key = "monthly:" + current + ":" + winner.ChatId + ":" + winner.Kind + ":" + winner.UserId;
                    if (Rows.ContainsKey(key)) continue;
                    Rows.Add(key, new RoomTitle { Key = key, ChatId = winner.ChatId, UserId = winner.UserId, Source = "monthly",
                        Title = winner.Kind == "monthly" ? "월간 활동왕" : "월간 인기왕", Reason = previous + " 공동 3위 이내",
                        GrantedAt = start, ExpiresAt = end });
                    Dirty = true;
                }
            }
        }
    }
}