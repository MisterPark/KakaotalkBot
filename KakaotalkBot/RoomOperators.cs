using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KakaotalkBot
{
    public sealed class RoomOperator
    {
        public long ChatId, UserId, LinkId, ObservedAt, LogId;
        public string Nickname, Privilege;
        public int MemberType;
        public bool IsPresent;
        public bool IsOperator { get { return IsPresent && (MemberType == 1 || MemberType == 4); } }
        public string Role { get { return MemberType == 1 ? "방장" : MemberType == 4 ? "부방장" : "운영진 아님"; } }
        internal List<object> ToRow()
        {
            return new List<object> { RoomUser.S(ChatId), RoomUser.S(UserId), Nickname ?? "", MemberType,
                Privilege ?? "", IsPresent, IsOperator, RoomUser.T(ObservedAt), RoomUser.S(LogId), RoomUser.S(LinkId) };
        }
    }

    // 저장된 행은 표시·보관용입니다. 명령 권한은 이번 수신 세션에서 읽은 자료로만 확인합니다.
    internal sealed class RoomOperatorStore
    {
        internal static readonly string[] Headers = { "chat_id", "user_id", "nickname", "member_type", "member_privilege", "is_present", "is_operator", "observed_at", "source_log_id", "link_id" };
        internal readonly Dictionary<Tuple<long, long>, RoomOperator> Records = new Dictionary<Tuple<long, long>, RoomOperator>();
        private readonly Dictionary<long, ChatRosterSnapshot> live = new Dictionary<long, ChatRosterSnapshot>();
        private readonly Dictionary<Tuple<long, long>, Tuple<long, int>> pending = new Dictionary<Tuple<long, long>, Tuple<long, int>>();
        private readonly HashSet<long> unknownChanges = new HashSet<long>();
        internal bool Dirty;
        internal void ResetLive() { live.Clear(); pending.Clear(); unknownChanges.Clear(); }

        internal void Load(List<List<string>> rows)
        {
            foreach (var r in rows)
            {
                long chat, user, log, link;
                int type;
                bool present, savedOperator;
                DateTimeOffset at;
                if (r.Count != Headers.Length || !Id(r[0], out chat) || !Id(r[1], out user) || !Id(r[9], out link) ||
                    !long.TryParse(r[8], NumberStyles.None, CultureInfo.InvariantCulture, out log) || log < 0 ||
                    !int.TryParse(r[3], out type) || !bool.TryParse(r[5], out present) || !bool.TryParse(r[6], out savedOperator) ||
                    !DateTimeOffset.TryParse(r[7], CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out at))
                    throw new FormatException("DB_RoomOperators의 행 형식이 잘못되었습니다.");
                var item = new RoomOperator { ChatId = chat, UserId = user, LinkId = link, LogId = log,
                    MemberType = type, Privilege = r[4], IsPresent = present, Nickname = r[2], ObservedAt = at.ToUnixTimeSeconds() };
                if (item.IsOperator != savedOperator) throw new FormatException("DB_RoomOperators의 운영진 여부와 역할이 일치하지 않습니다.");
                Records.Add(Tuple.Create(chat, user), item);
            }
            ResetLive(); Dirty = false;
        }
        private static bool Id(string s, out long id) { return long.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out id) && id > 0; }

        internal void Observe(ChatRosterSnapshot roster)
        {
            if (roster.ChatId <= 0 || !roster.IsOpenGroup) return;
            ChatRosterSnapshot previous;
            if (live.TryGetValue(roster.ChatId, out previous) && (roster.ObservedAt < previous.ObservedAt || roster.LogId < previous.LogId)) return;
            live[roster.ChatId] = roster;
            foreach (var member in roster.Members)
            {
                var key = Tuple.Create(roster.ChatId, member.UserId);
                Tuple<long, int> expected;
                if (pending.TryGetValue(key, out expected) && roster.LogId >= expected.Item1 &&
                    (member.MemberType == expected.Item2 || expected.Item2 == -2 && member.MemberType != 1 && member.MemberType >= 0 ||
                        expected.Item2 == -3 && !member.IsPresent || expected.Item2 == -4 && member.IsPresent))
                    pending.Remove(key);
                if (member.MemberType != 1 && member.MemberType != 4 && !Records.ContainsKey(key)) continue;
                Records[key] = new RoomOperator { ChatId = roster.ChatId, UserId = member.UserId, LinkId = roster.LinkId,
                    Nickname = member.Nickname, MemberType = member.MemberType, Privilege = member.MemberPrivilege,
                    IsPresent = member.IsPresent, ObservedAt = roster.ObservedAt, LogId = roster.LogId };
                Dirty = true;
            }
            foreach (var item in Records.Values.Where(r => r.ChatId == roster.ChatId && !roster.Members.Any(m => m.UserId == r.UserId)))
            { item.IsPresent = false; item.ObservedAt = roster.ObservedAt; item.LogId = roster.LogId; Dirty = true; }
            foreach (var pair in pending.Where(p => p.Key.Item1 == roster.ChatId && p.Value.Item2 == -3 &&
                roster.LogId >= p.Value.Item1 && !roster.Members.Any(m => m.UserId == p.Key.Item2)).ToArray()) pending.Remove(pair.Key);
        }

        internal void Invalidate(long chat, long user, long log, int expectedType)
        {
            if (user <= 0) { unknownChanges.Add(chat); return; }
            var key = Tuple.Create(chat, user);
            Tuple<long, int> old;
            if (!pending.TryGetValue(key, out old) || log >= old.Item1) pending[key] = Tuple.Create(log, expectedType);
        }

        internal bool Check(long chat, long user, long now, out string error)
        {
            ChatMemberState member;
            if (!TryMember(chat, user, now, out member, out error)) return false;
            if (member.MemberType != 1 && member.MemberType != 4)
            { error = "이 명령은 해당 방의 방장·부방장만 사용할 수 있습니다."; return false; }
            return true;
        }
        internal bool TryMember(long chat, long user, long now, out ChatMemberState member, out string error)
        {
            member = null;
            error = "운영진·참여자 정보를 확인 중입니다. 잠시 후 다시 시도해 주세요.";
            ChatRosterSnapshot roster;
            if (user <= 0 || !live.TryGetValue(chat, out roster) || unknownChanges.Contains(chat) ||
                pending.ContainsKey(Tuple.Create(chat, user)) || pending.Any(p => p.Key.Item1 == chat && p.Value.Item2 >= -2) ||
                now < roster.ObservedAt || now - roster.ObservedAt > 15) return false;
            member = roster.Members.FirstOrDefault(m => m.UserId == user && m.IsPresent);
            if (member == null || member.MemberType < 0) return false;
            error = null; return true;
        }
    }

    public static class OperatorCommandPolicy
    {
        private static readonly HashSet<string> names = new HashSet<string>(StringComparer.Ordinal) { "/이력", "/메모" };
        public static string Name(string text) { return (text ?? "").TrimStart().Split((char[])null, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? ""; }
        public static bool RequiresOperator(string text) { return names.Contains(Name(text)); }
        public static void Register(string command)
        {
            if (string.IsNullOrWhiteSpace(command) || !command.StartsWith("/", StringComparison.Ordinal) || Name(command) != command)
                throw new ArgumentException("공백 없는 명령어 이름을 입력해 주세요.", "command");
            names.Add(command);
        }
    }
}
