using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KakaotalkBot
{
    public sealed class RoomUser
    {
        public long ChatId, UserId, PresenceLogId, LastMessageLogId, LastExperienceAt, NicknameAt;
        public string Nickname;
        public bool? IsPresent;
        public long LastSeenAt;
        public List<object> ToRow()
        {
            return new List<object> { S(ChatId), S(UserId), Nickname ?? "", IsPresent.HasValue ? (object)IsPresent.Value : "",
                T(LastSeenAt), S(PresenceLogId), S(LastMessageLogId), T(LastExperienceAt), T(NicknameAt) };
        }
        internal static string S(long value) { return value.ToString(CultureInfo.InvariantCulture); }
        internal static string T(long value) { return value == 0 ? "" : DateTimeOffset.FromUnixTimeSeconds(value).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture); }
    }

    public sealed class RoomEventRecord
    {
        public string Key, Kind, Nickname;
        public long ChatId, UserId, LogId, OccurredAt;
        public List<object> ToRow() { return new List<object> { Key, RoomUser.S(ChatId), RoomUser.S(UserId), Kind,
            RoomUser.T(OccurredAt), RoomUser.S(LogId), Nickname ?? "" }; }
    }

    public sealed class NicknameRecord
    {
        public string Key, Before, After, Source;
        public long ChatId, UserId, ObservedAt, LogId;
        public List<object> ToRow() { return new List<object> { Key, RoomUser.S(ChatId), RoomUser.S(UserId), Before, After,
            RoomUser.T(ObservedAt), RoomUser.S(LogId), Source }; }
    }

    // 방 상태와 이력은 계정의 전역 포인트/경험치와 분리합니다.
    internal sealed class UserActivityStore
    {
        internal static readonly string[] RoomHeaders = { "chat_id", "user_id", "nickname", "is_present", "last_seen_at", "presence_log_id", "last_message_log_id", "last_experience_at", "nickname_observed_at" };
        internal static readonly string[] EventHeaders = { "event_key", "chat_id", "user_id", "event_type", "occurred_at", "log_id", "nickname" };
        internal static readonly string[] NicknameHeaders = { "event_key", "chat_id", "user_id", "old_nickname", "new_nickname", "observed_at", "log_id", "source" };
        internal readonly Dictionary<Tuple<long, long>, RoomUser> Rooms = new Dictionary<Tuple<long, long>, RoomUser>();
        internal readonly List<RoomEventRecord> Events = new List<RoomEventRecord>();
        internal readonly List<NicknameRecord> Nicknames = new List<NicknameRecord>();
        private readonly HashSet<string> eventKeys = new HashSet<string>();
        internal int SavedEvents, SavedNicknames;
        internal bool Dirty;

        internal RoomUser Get(long chat, long user)
        {
            if (chat <= 0 || user <= 0) throw new ArgumentOutOfRangeException("user");
            RoomUser result;
            var key = Tuple.Create(chat, user);
            if (!Rooms.TryGetValue(key, out result))
            {
                result = new RoomUser { ChatId = chat, UserId = user };
                Rooms.Add(key, result); Dirty = true;
            }
            return result;
        }
        internal static long Time(long seconds) { return seconds > 0 ? seconds : DateTimeOffset.UtcNow.ToUnixTimeSeconds(); }
        internal static string EventKey(long chat, long user, long log) { return chat + ":" + log + ":" + user; }
        internal bool HasEvent(long chat, long user, long log) { return eventKeys.Contains(EventKey(chat, user, log)); }

        internal bool RecordEvent(long chat, long user, long log, string kind, string name, long time)
        {
            string key = EventKey(chat, user, log);
            if (log <= 0 || !eventKeys.Add(key)) return false;
            time = Time(time);
            Events.Add(new RoomEventRecord { Key = key, ChatId = chat, UserId = user, LogId = log, Kind = kind, Nickname = name, OccurredAt = time });
            Observe(chat, user, name, kind == "join", time, log, "event");
            Dirty = true;
            return true;
        }

        internal bool Observe(long chat, long user, string name, bool present, long time, long log, string source)
        {
            var room = Get(chat, user);
            time = Time(time);
            bool changedName = false;
            if (!string.IsNullOrWhiteSpace(name) && time >= room.NicknameAt && log >= room.PresenceLogId)
            {
                if (room.Nickname != name)
                {
                    // 최초 확인한 이름을 기준으로 삼고 이전 이름을 추측해 이력을 만들지 않습니다.
                    if (!string.IsNullOrWhiteSpace(room.Nickname))
                    {
                        Nicknames.Add(new NicknameRecord { Key = Guid.NewGuid().ToString("N"), ChatId = chat, UserId = user,
                            Before = room.Nickname, After = name, ObservedAt = time, LogId = log, Source = source });
                        changedName = true;
                    }
                    room.Nickname = name; Dirty = true;
                }
                if (time > room.NicknameAt) { room.NicknameAt = time; Dirty = true; }
            }
            // 늦게 도착한 과거 퇴장으로 재입장 후 상태를 덮어쓰지 않습니다.
            if (log >= room.PresenceLogId)
            {
                if (room.IsPresent != present || room.PresenceLogId != log) Dirty = true;
                room.IsPresent = present; room.PresenceLogId = log;
            }
            if (present && time > room.LastSeenAt) { room.LastSeenAt = time; Dirty = true; }
            return changedName;
        }

        internal bool AwardExperience(User user, long chat, long log, long at)
        {
            if (at <= 0 || log <= 0) return false;
            var room = Get(chat, user.UserId);
            if (log <= room.LastMessageLogId) return false;
            room.LastMessageLogId = log; Dirty = true;
            long last = Rooms.Values.Where(r => r.UserId == user.UserId).Max(r => r.LastExperienceAt);
            if (at < last || (last > 0 && at - last < 60)) return false;
            user.Experience = checked(user.Experience + 1);
            room.LastExperienceAt = at;
            return true;
        }

        internal void Load(List<List<string>> rooms, List<List<string>> events, List<List<string>> names)
        {
            foreach (var r in rooms)
            {
                var room = new RoomUser { ChatId = N(r, 0, true), UserId = N(r, 1, true), Nickname = r[2],
                    IsPresent = string.IsNullOrWhiteSpace(r[3]) ? (bool?)null : bool.Parse(r[3]), LastSeenAt = At(r, 4),
                    PresenceLogId = N(r, 5), LastMessageLogId = N(r, 6), LastExperienceAt = At(r, 7), NicknameAt = At(r, 8) };
                Rooms.Add(Tuple.Create(room.ChatId, room.UserId), room);
            }
            foreach (var r in events)
            {
                var item = new RoomEventRecord { Key = r[0], ChatId = N(r, 1, true), UserId = N(r, 2, true), Kind = r[3], OccurredAt = At(r, 4), LogId = N(r, 5, true), Nickname = r[6] };
                if (item.Key != EventKey(item.ChatId, item.UserId, item.LogId) || !eventKeys.Add(item.Key) || !new[] { "join", "leave", "kick" }.Contains(item.Kind))
                    throw new FormatException("입퇴장 이력의 키 또는 종류가 잘못되었습니다.");
                Events.Add(item);
            }
            var nameKeys = new HashSet<string>();
            foreach (var r in names)
            {
                if (string.IsNullOrWhiteSpace(r[0]) || !nameKeys.Add(r[0])) throw new FormatException("닉네임 이력 키가 중복되거나 비어 있습니다.");
                Nicknames.Add(new NicknameRecord { Key = r[0], ChatId = N(r, 1, true), UserId = N(r, 2, true), Before = r[3], After = r[4], ObservedAt = At(r, 5), LogId = N(r, 6), Source = r[7] });
            }
            MarkSaved();
        }
        private static long N(List<string> row, int column, bool positive = false)
        {
            long value;
            if (!long.TryParse(row[column], NumberStyles.None, CultureInfo.InvariantCulture, out value) || value < (positive ? 1 : 0))
                throw new FormatException("사용자 활동 DB의 ID 또는 시각이 잘못되었습니다.");
            return value;
        }
        private static long At(List<string> row, int column)
        {
            if (string.IsNullOrWhiteSpace(row[column])) return 0;
            DateTimeOffset value;
            if (!DateTimeOffset.TryParse(row[column], CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out value) || value.ToUnixTimeSeconds() < 0)
                throw new FormatException("활동 DB의 시각은 시간대가 포함된 ISO 8601 형식이어야 합니다.");
            return value.ToUnixTimeSeconds();
        }
        internal void MarkSaved() { SavedEvents = Events.Count; SavedNicknames = Nicknames.Count; Dirty = false; }
    }
}
