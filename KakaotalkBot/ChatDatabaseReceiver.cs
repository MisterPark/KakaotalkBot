using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace KakaotalkBot
{
    public sealed class ChatRoomInfo
    {
        public string AccountPath { get; internal set; }
        public int ProcessId { get; internal set; }
        public long ChatId { get; internal set; }
        public long LinkId { get; internal set; }
        public long DirectUserId { get; internal set; }
        public string Name { get; internal set; }
        public string Kind { get; internal set; }
        public string SourcePath { get { return Path.Combine(AccountPath, "chat_data", "chatLogs_" + ChatId + ".edb"); } }
        public bool IsDirect { get { return Kind == "DirectChat" || Kind == "OD"; } }
        public bool IsOpenGroup { get { return LinkId > 0 && !IsDirect; } }
        public string Identity { get { return AccountPath + ":" + ChatId; } }
    }

    internal enum ChatRoomEvent { None, Join, Leave, Kick }

    internal sealed class ChatMessage
    {
        public ChatRosterSnapshot Roster;
        public long ChatId, LogId, AuthorId, SendAt;
        public string Nickname, Message, EventCommand;
        public bool IsDirect;
        public bool IsOwn;
        public bool RoleChanged;
        public int ExpectedMemberType;
        public ChatRoomEvent RoomEvent;
        public IReadOnlyList<ChatMention> Mentions = ChatMention.Empty;
    }

    internal sealed class ChatMemberState
    {
        public long UserId;
        public string Nickname;
        public bool IsPresent;
        public int MemberType = -1;
        public string MemberPrivilege;
    }
    internal sealed class ChatRosterSnapshot
    {
        public long ChatId, LogId, ObservedAt;
        public long LinkId;
        public bool IsOpenGroup;
        public readonly List<ChatMemberState> Members = new List<ChatMemberState>();
        public static ChatRosterSnapshot Read(string path, ChatRoomInfo room, ChatUserDirectory users)
        {
            var result = new ChatRosterSnapshot { ChatId = room.ChatId, LinkId = room.LinkId, IsOpenGroup = room.IsOpenGroup,
                ObservedAt = DateTimeOffset.UtcNow.ToUnixTimeSeconds() };
            using (var db = new ChatSqlite(path))
            {
                result.LogId = db.Scalar("SELECT COALESCE(lastLogId,0) FROM chatRoomList WHERE chatId=?", room.ChatId);
                foreach (var row in db.Query("SELECT userId,isActive FROM chatMembers WHERE chatId=? AND userId>0 AND isActive IN (0,1)", room.ChatId))
                    result.Members.Add(new ChatMemberState { UserId = ChatCatalog.Number(row[0]), IsPresent = ChatCatalog.Number(row[1]) == 1,
                        Nickname = users.Find(ChatCatalog.Number(row[0]), room.LinkId),
                        MemberType = users.MemberType(ChatCatalog.Number(row[0]), room.LinkId),
                        MemberPrivilege = users.MemberPrivilege(ChatCatalog.Number(row[0]), room.LinkId) });
            }
            return result;
        }
    }

    internal sealed class ChatCatalog
    {
        public readonly List<ChatRoomInfo> Rooms = new List<ChatRoomInfo>();
        public readonly List<string> Errors = new List<string>();
        public static string OutputRoot { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DecryptedChat"); } }

        public static ChatCatalog Load(CancellationToken cancellation)
        {
            var result = new ChatCatalog();
            var decryptor = new KakaoTalkDecryptor();
            var directories = KakaoTalkDecryptor.FindChatDataDirectories();
            var targets = new List<KakaoTalkDecryptor.DatabaseInfo>();
            foreach (string directory in directories)
            foreach (string path in new[] { Path.Combine(directory, "chatListInfo.edb"), Path.Combine(Path.GetDirectoryName(directory), "TalkUserDB.edb") })
            {
                if (File.Exists(path)) targets.Add(decryptor.ReadDatabaseInfo(path));
            }
            int[] processes = KakaoTalkDecryptor.FindKakaoTalkProcessIds();
            if (processes.Length == 0) { result.Errors.Add("카카오톡을 실행한 뒤 방 목록을 다시 읽어 주세요."); return result; }
            bool foundListKey = false;
            var handled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (int pid in processes)
            {
                cancellation.ThrowIfCancellationRequested();
                var keys = decryptor.DiscoverKeys(pid, targets, cancellation);
                try
                {
                    foreach (string directory in directories)
                    {
                        string account = Path.GetDirectoryName(directory);
                        if (handled.Contains(account)) continue;
                        string listPath = Path.Combine(directory, "chatListInfo.edb");
                        KakaoTalkDecryptor.DecryptionKey listKey;
                        if (!keys.TryGetValue(listPath, out listKey)) continue;
                        foundListKey = true;
                        string output = Path.Combine(OutputRoot, Path.GetFileName(account));
                        Directory.CreateDirectory(output);
                        using (var lease = new FileStream(Path.Combine(output, "catalog.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                        {
                            var list = new ChatDatabaseSnapshot(listPath, Path.Combine(output, "catalog-rooms.sqlite"), decryptor, listKey);
                            if (!RefreshCatalogSnapshot(list, cancellation)) { result.Errors.Add("방 목록 DB 갱신 중입니다. 잠시 후 다시 읽어 주세요."); continue; }
                            string usersPath = Path.Combine(account, "TalkUserDB.edb");
                            var users = new ChatUserDirectory();
                            KakaoTalkDecryptor.DecryptionKey usersKey;
                            if (keys.TryGetValue(usersPath, out usersKey))
                            {
                                var usersDb = new ChatDatabaseSnapshot(usersPath, Path.Combine(output, "catalog-users.sqlite"), decryptor, usersKey);
                                if (!RefreshCatalogSnapshot(usersDb, cancellation)) { result.Errors.Add("사용자 DB 갱신 중입니다. 잠시 후 다시 읽어 주세요."); continue; }
                                users.Load(usersDb.OutputPath, list.OutputPath);
                            }
                            result.Rooms.AddRange(ReadRooms(list.OutputPath, account, pid, users));
                            handled.Add(account);
                        }
                    }
                }
                finally { foreach (var key in keys.Values) key.Dispose(); }
            }
            if (result.Rooms.Count == 0 && result.Errors.Count == 0)
                result.Errors.Add(targets.Count == 0 ? "계정 DB 파일을 찾지 못했습니다. 카카오톡 로그인과 데이터 폴더를 확인하세요." :
                    foundListKey ? "계정 DB는 읽었지만 채팅방 목록이 비어 있습니다." :
                    "계정 DB의 키를 찾지 못했습니다. 카카오톡에 로그인하고 채팅 목록을 연 뒤 다시 시도하세요.");
            return result;
        }

        // 방 목록은 주기 수신과 달리 일회성 작업이므로 같은 키로 잠시 재시도합니다.
        private static bool RefreshCatalogSnapshot(ChatDatabaseSnapshot snapshot, CancellationToken cancellation)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                cancellation.ThrowIfCancellationRequested();
                snapshot.Refresh(cancellation);
                if (!snapshot.RefreshDeferred) return true;
                if (attempt < 4 && cancellation.WaitHandle.WaitOne(200)) cancellation.ThrowIfCancellationRequested();
            }
            return false;
        }

        internal static List<ChatRoomInfo> ReadRooms(string path, string account, int pid, ChatUserDirectory users)
        {
            using (var db = new ChatSqlite(path))
                return db.Query("SELECT chatId, chatRoomTitle, type, directChatMemberId, linkId FROM chatRoomList ORDER BY lastUpdatedAt DESC")
                    .Select(row => new ChatRoomInfo
                    {
                        AccountPath = account, ProcessId = pid, ChatId = Number(row[0]), Kind = Convert.ToString(row[2]),
                        DirectUserId = Number(row[3]), LinkId = Number(row[4]),
                        Name = !string.IsNullOrWhiteSpace(Convert.ToString(row[1])) ? Convert.ToString(row[1])
                            : users.Find(Number(row[3]), Number(row[4])) ?? ("이름 없는 방 " + Number(row[0]))
                    }).ToList();
        }

        internal static long Number(object value) { return value == null ? 0 : Convert.ToInt64(value); }
    }

    internal sealed class ChatUserDirectory
    {
        private readonly Dictionary<string, string> names = new Dictionary<string, string>();
        private readonly Dictionary<string, int> roles = new Dictionary<string, int>();
        private readonly Dictionary<string, string> privileges = new Dictionary<string, string>();
        public readonly HashSet<long> OwnIds = new HashSet<long>();

        public void Load(string usersPath, string roomsPath)
        {
            var next = new Dictionary<string, string>();
            var nextRoles = new Dictionary<string, int>();
            var nextPrivileges = new Dictionary<string, string>();
            using (var db = new ChatSqlite(usersPath))
            {
            var columns = new HashSet<string>(db.Query("PRAGMA table_info(talkUser)").Select(r => Convert.ToString(r[1])));
            string role = columns.Contains("openMemberType") ? "openMemberType" : "NULL";
            string privilege = columns.Contains("openMemberPrivilege") ? "CAST(openMemberPrivilege AS TEXT)" : "NULL";
            foreach (object[] row in db.Query("SELECT userId, linkId, nickName, friendNickName," + role + "," + privilege + " FROM talkUser"))
            {
                long id = ChatCatalog.Number(row[0]), link = ChatCatalog.Number(row[1]);
                string name = Convert.ToString(link == 0 && !string.IsNullOrWhiteSpace(Convert.ToString(row[3])) ? row[3] : row[2]);
                if (!string.IsNullOrWhiteSpace(name)) next[id + ":" + link] = name;
                nextRoles[id + ":" + link] = row[4] == null ? -1 : checked((int)ChatCatalog.Number(row[4]));
                nextPrivileges[id + ":" + link] = Convert.ToString(row[5], CultureInfo.InvariantCulture);
            }
            }
            names.Clear();
            roles.Clear(); privileges.Clear();
            foreach (var pair in nextRoles) roles[pair.Key] = pair.Value;
            foreach (var pair in nextPrivileges) privileges[pair.Key] = pair.Value;
            foreach (var pair in next) names[pair.Key] = pair.Value;
            using (var db = new ChatSqlite(roomsPath))
            foreach (var row in db.Query("SELECT DISTINCT m.userId FROM chatMembers m JOIN chatRoomList r ON r.chatId=m.chatId WHERE r.type='MemoChat'"))
                OwnIds.Add(ChatCatalog.Number(row[0]));
        }

        public string Find(long author, long link)
        {
            string name;
            if (names.TryGetValue(author + ":" + link, out name)) return name;
            // 오픈채팅 닉네임을 일반 프로필 이름으로 대체하지 않습니다.
            return null;
        }
        public int MemberType(long user, long link)
        { int value; return link > 0 && roles.TryGetValue(user + ":" + link, out value) ? value : -1; }
        public string MemberPrivilege(long user, long link)
        { string value; return privileges.TryGetValue(user + ":" + link, out value) ? value : ""; }
    }

    internal sealed class ChatLogCursor
    {
        private readonly HashSet<long> seen = new HashSet<long>();
        private readonly Dictionary<long, string> systemContents = new Dictionary<long, string>();
        private long floor, last;
        private bool initialized;
        private readonly long since;
        public ChatLogCursor(long since = 0) { this.since = since; }

        public List<object[]> Read(string path)
        {
            using (var db = new ChatSqlite(path))
            {
                if (!initialized)
                {
                    floor = since == 0 ? db.Scalar("SELECT COALESCE(MAX(logId),0) FROM chatLogs") :
                        Math.Max(0, db.Scalar("SELECT COALESCE(MIN(logId),1) FROM chatLogs WHERE sendAt>=?", since) - 1);
                    last = floor;
                    initialized = true;
                    if (since == 0) return new List<object[]>();
                }
                var rows = db.Query("SELECT logId,authorId,type,sendAt,message,COALESCE(deleted,0),COALESCE(write_on_pc,0),attachement FROM chatLogs WHERE logId>? AND sendAt>=? ORDER BY logId LIMIT 256", last, since);
                if (rows.Count == 0)
                {
                    // 최근 구간을 다시 읽어 지연 저장된 메시지를 보완합니다. 시작 이전 기록은 제외합니다.
                    rows = db.Query("SELECT logId,authorId,type,sendAt,message,COALESCE(deleted,0),COALESCE(write_on_pc,0),attachement FROM chatLogs WHERE logId>? AND sendAt>=? ORDER BY logId DESC LIMIT 256", floor, since);
                    rows.Reverse();
                }
                // 시스템 메시지는 행이 먼저 생기고 내용이 나중에 확정되거나 이전 로그 번호로 들어올 수 있습니다.
                // 새 일반 메시지가 계속 들어와도 최근 시스템 행을 재확인합니다. 시작 이전 기록은 제외합니다.
                var systemRows = db.Query("SELECT logId,authorId,type,sendAt,message,COALESCE(deleted,0),COALESCE(write_on_pc,0),attachement FROM chatLogs WHERE logId>? AND sendAt>=? AND type=0 ORDER BY logId DESC LIMIT 256", floor, since);
                return rows.Concat(systemRows).GroupBy(row => ChatCatalog.Number(row[0])).Select(group => group.Last())
                    .Where(row =>
                    {
                        long id = ChatCatalog.Number(row[0]);
                        if (!seen.Contains(id)) return true;
                        if (ChatCatalog.Number(row[2]) != 0) return false;
                        string previous;
                        return !systemContents.TryGetValue(id, out previous) || previous != SystemContent(row);
                    }).OrderBy(row => ChatCatalog.Number(row[0])).ToList();
            }
        }

        private static string SystemContent(object[] row)
        { return Convert.ToString(row[5]) + ":" + Convert.ToString(row[4]); }

        public void Accept(object[] row)
        {
            long id = ChatCatalog.Number(row[0]);
            if (ChatCatalog.Number(row[2]) == 0) systemContents[id] = SystemContent(row);
            Accept(id);
        }

        public void Accept(long id)
        {
            seen.Add(id);
            last = Math.Max(last, id);
            if (seen.Count > 4096)
            {
                long boundary = seen.OrderByDescending(x => x).Skip(2047).First();
                seen.RemoveWhere(x => x < boundary);
                foreach (long old in systemContents.Keys.Where(x => x < boundary).ToArray()) systemContents.Remove(old);
                floor = Math.Max(floor, boundary - 1);
            }
        }
    }

    internal static class ChatMessageDecoder
    {
        public static List<ChatMessage> Decode(ChatRoomInfo room, object[] row, ChatUserDirectory users, bool direct)
        {
            var result = new List<ChatMessage>();
            long author = ChatCatalog.Number(row[1]);
            int type = (int)ChatCatalog.Number(row[2]);
            if (author > 0 && (type == 1 || type == 26) && ChatCatalog.Number(row[6]) != 0) users.OwnIds.Add(author);
            if (ChatCatalog.Number(row[5]) != 0) return result;
            if (type == 1 || type == 26)
            {
                if (author <= 0 || direct && users.OwnIds.Contains(author)) return result;
                string nickname = users.Find(author, room.LinkId);
                // 프로필 반영이 늦어도 ID로 수신을 계속합니다. 이름은 후속 프로필 관찰로 보완합니다.
                var message = Message(room, row, author, nickname, Convert.ToString(row[4]), null, direct);
                message.IsOwn = users.OwnIds.Contains(author);
                message.Mentions = ReadMentions(row.Length > 7 ? Convert.ToString(row[7]) : null, room.LinkId, users);
                result.Add(message);
            }
            else if (type == 0 && !direct)
            {
                JObject feed;
                try { feed = JObject.Parse(Convert.ToString(row[4])); }
                catch (Newtonsoft.Json.JsonException) { return result; }
                long feedType;
                if (!TryPositiveNumber(feed["feedType"], out feedType)) return result;
                if (feedType == 11 || feedType == 12 || feedType == 15)
                {
                    string[] fields = feedType == 15 ? new[] { "prevHost", "newHost" } : new[] { "member" };
                    foreach (string field in fields)
                    {
                        long id;
                        var member = feed[field] as JObject;
                        if (member == null || !TryPositiveNumber(member["userId"], out id)) id = 0;
                        var changed = Message(room, row, id, null, null, null, false);
                        changed.RoleChanged = true;
                        changed.ExpectedMemberType = field == "prevHost" ? -2 : field == "newHost" ? 1 : feedType == 11 ? 4 : 2;
                        result.Add(changed);
                    }
                    return result;
                }
                if ((bool?)feed["hidden"] == true) return result;
                var roomEvent = feedType == 1 || feedType == 4 ? ChatRoomEvent.Join :
                    feedType == 2 ? ChatRoomEvent.Leave : feedType == 6 ? ChatRoomEvent.Kick : ChatRoomEvent.None;
                if (roomEvent == ChatRoomEvent.None) return result;
                // 기존 입퇴장 응답은 유지하면서, 집계용 원본 이벤트 종류를 별도로 전달한다.
                string command = roomEvent == ChatRoomEvent.Join ? "/입장" : "/퇴장";
                var members = feed["members"] as JArray ?? new JArray();
                if (feed["member"] is JObject) members.Add(feed["member"]);
                var handled = new HashSet<long>();
                foreach (JObject member in members.OfType<JObject>())
                {
                    long id;
                    if (!TryPositiveNumber(member["userId"], out id) || !handled.Add(id)) continue;
                    string name = member["nickName"] != null && member["nickName"].Type == JTokenType.String ? (string)member["nickName"] : null;
                    if (string.IsNullOrWhiteSpace(name)) name = users.Find(id, room.LinkId);
                    var message = Message(room, row, id, name, command, command, false);
                    message.RoomEvent = roomEvent;
                    message.IsOwn = users.OwnIds.Contains(id);
                    result.Add(message);
                }
            }
            return result;
        }

        private static IReadOnlyList<ChatMention> ReadMentions(string attachment, long linkId, ChatUserDirectory users)
        {
            if (string.IsNullOrWhiteSpace(attachment)) return ChatMention.Empty;
            JObject metadata;
            try { metadata = JObject.Parse(attachment); }
            catch (Newtonsoft.Json.JsonException) { return ChatMention.Empty; }

            // 답장이 인용한 원본의 src_mentions는 현재 메시지의 태그 대상으로 사용하지 않습니다.
            var entries = metadata["mentions"] as JArray;
            if (entries == null) return ChatMention.Empty;
            var result = new List<ChatMention>();
            foreach (JObject entry in entries.OfType<JObject>())
            {
                long userId, length;
                var at = entry["at"] as JArray;
                if (!TryPositiveNumber(entry["user_id"], out userId) ||
                    !TryPositiveNumber(entry["len"], out length) || length > int.MaxValue || at == null) continue;
                var positions = new List<int>();
                foreach (JToken index in at)
                {
                    long position;
                    if (TryPositiveNumber(index, out position) && position <= int.MaxValue) positions.Add((int)position);
                }
                if (positions.Count == 0) continue;
                // 프로필 갱신이 늦더라도 확인된 사용자 ID는 바로 전달합니다.
                result.Add(new ChatMention(userId, positions, (int)length, users.Find(userId, linkId)));
            }
            return Array.AsReadOnly(result.OrderBy(mention => mention.At[0]).ToArray());
        }

        private static bool TryPositiveNumber(JToken token, out long value)
        {
            value = 0;
            if (token == null || (token.Type != JTokenType.Integer && token.Type != JTokenType.String)) return false;
            string text = token.Type == JTokenType.String ? (string)token : token.ToString(Newtonsoft.Json.Formatting.None);
            return long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) && value > 0;
        }

        private static ChatMessage Message(ChatRoomInfo room, object[] row, long author, string name, string text, string command, bool direct)
        {
            return new ChatMessage { ChatId = room.ChatId, LogId = ChatCatalog.Number(row[0]), AuthorId = author,
                SendAt = ChatCatalog.Number(row[3]), Nickname = name, Message = text, EventCommand = command, IsDirect = direct };
        }
    }

    // 집계 누락을 추적하기 위한 최소 기록입니다. 채팅 본문이나 닉네임은 남기지 않습니다.
    internal static class ChatEventDiagnostics
    {
        private static readonly object Gate = new object();
        public static void Write(string stage, ChatMessage message = null)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(ChatCatalog.OutputRoot);
                    string path = Path.Combine(ChatCatalog.OutputRoot, "departure-events.log");
                    if (File.Exists(path) && new FileInfo(path).Length > 1024 * 1024)
                    {
                        File.Copy(path, path + ".previous", true);
                        File.WriteAllText(path, "");
                    }
                    string detail = message == null ? "" : string.Format(CultureInfo.InvariantCulture,
                        " room={0} log={1} user={2} event={3}", message.ChatId, message.LogId, message.AuthorId, message.RoomEvent);
                    File.AppendAllText(path, DateTimeOffset.Now.ToString("o", CultureInfo.InvariantCulture) + " " + stage + detail + "\r\n");
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    internal sealed class ChatDatabaseReceiver : IDisposable
    {
        private readonly ChatRoomInfo target;
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly BlockingCollection<ChatMessage> messages = new BlockingCollection<ChatMessage>(4096);
        private readonly ConcurrentQueue<DirectRequest> requests = new ConcurrentQueue<DirectRequest>();
        private readonly Dictionary<string, Session> sessions = new Dictionary<string, Session>(StringComparer.OrdinalIgnoreCase);
        private readonly KakaoTalkDecryptor decryptor = new KakaoTalkDecryptor();
        private readonly ChatUserDirectory users = new ChatUserDirectory();
        private readonly long ownId;
        private Task worker;
        private int disposed;
        private volatile string status = "준비 중";
        public string Status { get { return status; } }
        public Task Completion { get { return worker ?? Task.FromResult(0); } }

        public ChatDatabaseReceiver(ChatRoomInfo room, long ownId = 0) { target = room; this.ownId = ownId; }
        public void Start() { if (worker == null) worker = Task.Run((Action)Run); }
        public bool TryRead(out ChatMessage message) { return messages.TryTake(out message); }
        public void WatchDirect(long authorId)
        {
            if (authorId <= 0) throw new ArgumentOutOfRangeException("authorId");
            requests.Enqueue(new DirectRequest { AuthorId = authorId,
                Since = DateTimeOffset.UtcNow.ToUnixTimeSeconds() });
        }

        internal static ChatRoomInfo FindDirectRoom(IEnumerable<ChatRoomInfo> rooms, long authorId, long sourceChatId)
        {
            if (authorId <= 0) return null;
            var matches = rooms.Where(room => room.IsDirect && room.ChatId != sourceChatId && room.DirectUserId == authorId).Take(2).ToList();
            return matches.Count == 1 ? matches[0] : null;
        }

        private void Run()
        {
            string output = Path.Combine(ChatCatalog.OutputRoot, Path.GetFileName(target.AccountPath));
            var direct = new Dictionary<long, DirectRequest>();
            var rooms = new List<ChatRoomInfo>();
            DateTime nextMetadata = DateTime.MinValue;
            DateTime nextKeys = DateTime.MinValue;
            try
            {
                Directory.CreateDirectory(output);
                using (var lease = new FileStream(Path.Combine(output, "receiver.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    if (ownId > 0) users.OwnIds.Add(ownId);
                    while (!cancellation.IsCancellationRequested)
                    {
                        try
                        {
                            DirectRequest request;
                            while (requests.TryDequeue(out request)) { direct[request.AuthorId] = request; nextMetadata = DateTime.MinValue; }
                            foreach (long id in direct.Where(p => DateTimeOffset.UtcNow.ToUnixTimeSeconds() - p.Value.Since > 300).Select(p => p.Key).ToArray()) direct.Remove(id);
                            string listSource = Path.Combine(target.AccountPath, "chat_data", "chatListInfo.edb");
                            string userSource = Path.Combine(target.AccountPath, "TalkUserDB.edb");
                            var paths = new List<string> { listSource, userSource, target.SourcePath };
                            var directRooms = new Dictionary<long, DirectRequest>();
                            foreach (var watch in direct.Values)
                            {
                                var match = FindDirectRoom(rooms, watch.AuthorId, target.ChatId);
                                if (match != null) { paths.Add(match.SourcePath); directRooms[match.ChatId] = watch; }
                            }
                            foreach (string path in sessions.Keys.Where(p => !paths.Contains(p, StringComparer.OrdinalIgnoreCase)).ToArray())
                            { sessions[path].Dispose(); sessions.Remove(path); }
                            if (DateTime.UtcNow >= nextKeys)
                            {
                                nextKeys = DateTime.UtcNow.AddSeconds(5);
                                var missing = paths.Distinct(StringComparer.OrdinalIgnoreCase).Where(p => (!sessions.ContainsKey(p) || sessions[p].Key == null) && File.Exists(p)).Select(decryptor.ReadDatabaseInfo).ToList();
                                if (missing.Count > 0)
                                {
                                    status = "채팅 DB 키를 찾는 중";
                                    foreach (int pid in KakaoTalkDecryptor.FindKakaoTalkProcessIds().OrderBy(p => p == target.ProcessId ? 0 : 1))
                                    {
                                        var keys = decryptor.DiscoverKeys(pid, missing, cancellation.Token);
                                        foreach (var pair in keys)
                                        {
                                            Session previous;
                                            if (sessions.TryGetValue(pair.Key, out previous) && previous.Key != null) { pair.Value.Dispose(); continue; }
                                            var replacement = new Session(pair.Value, new ChatDatabaseSnapshot(pair.Key,
                                                Path.Combine(output, Path.GetFileNameWithoutExtension(pair.Key) + ".sqlite"), decryptor, pair.Value));
                                            if (previous != null) { replacement.Cursor = previous.Cursor; replacement.Since = previous.Since; }
                                            sessions[pair.Key] = replacement;
                                        }
                                        missing.RemoveAll(d => sessions.ContainsKey(d.SourcePath) && sessions[d.SourcePath].Key != null);
                                        if (missing.Count == 0) break;
                                    }
                                }
                            }
                            Session list, user;
                            if (!sessions.TryGetValue(listSource, out list) || list.Key == null || !sessions.TryGetValue(userSource, out user) || user.Key == null)
                                throw new IOException("계정/사용자 DB의 키를 기다립니다. 카카오톡의 채팅 목록과 대상 방을 열어 주세요.");
                            if (DateTime.UtcNow >= nextMetadata)
                            {
                                list.Snapshot.Refresh(cancellation.Token);
                                user.Snapshot.Refresh(cancellation.Token);
                                if (list.Snapshot.RefreshDeferred || user.Snapshot.RefreshDeferred)
                                { status = "계정 DB 갱신 중 · 다음 수신 주기에 재시도"; cancellation.Token.WaitHandle.WaitOne(200); continue; }
                                users.Load(user.Snapshot.OutputPath, list.Snapshot.OutputPath);
                                rooms = ChatCatalog.ReadRooms(list.Snapshot.OutputPath, target.AccountPath, target.ProcessId, users);
                                messages.Add(new ChatMessage { ChatId = target.ChatId,
                                    Roster = ChatRosterSnapshot.Read(list.Snapshot.OutputPath, target, users) }, cancellation.Token);
                                nextMetadata = DateTime.UtcNow.AddSeconds(5);
                            }
                            Session main;
                            if (!sessions.TryGetValue(target.SourcePath, out main) || main.Key == null) throw new IOException("대상 방의 키를 기다립니다. 카카오톡에서 해당 채팅방을 열어 주세요.");
                            Pump(main, target, null);
                            if (main.Snapshot.RefreshDeferred)
                            { status = "채팅 DB 갱신 중 · 다음 수신 주기에 재시도"; cancellation.Token.WaitHandle.WaitOne(200); continue; }
                            string directStatus = "";
                            foreach (var pair in directRooms)
                            {
                                var room = rooms.First(r => r.ChatId == pair.Key);
                                Session session;
                                if (sessions.TryGetValue(room.SourcePath, out session) && session.Key != null)
                                {
                                    try { Pump(session, room, pair.Value); }
                                    catch (IOException) { directStatus = " · 1:1 DB 읽기 대기"; }
                                }
                                else directStatus = " · 1:1 방을 열어 주세요";
                            }
                            if (direct.Count > 0 && directRooms.Count == 0) directStatus = " · 1:1 방 목록 대기";
                            status = "수신 중 · 최근 복호화 " + main.Snapshot.LastDecryptedPages + "페이지" + directStatus;
                        }
                        catch (CryptographicException)
                        {
                            status = "DB 키가 변경되었거나 페이지 검증에 실패했습니다. 키를 다시 확인합니다.";
                            foreach (var session in sessions.Values) session.Dispose();
                            // 읽기 기준점은 Session에 보관하므로 키 재발견 시에도 이어서 읽습니다.
                            foreach (var session in sessions.Values) session.Key = null;
                            nextMetadata = DateTime.MinValue;
                            nextKeys = DateTime.UtcNow.AddSeconds(5);
                            cancellation.Token.WaitHandle.WaitOne(1000);
                        }
                        catch (OperationCanceledException) { break; }
                        catch (Exception ex)
                        {
                            status = "수신 대기: " + ex.Message;
                            cancellation.Token.WaitHandle.WaitOne(1000);
                        }
                        cancellation.Token.WaitHandle.WaitOne(200);
                    }
                }
            }
            catch (Exception ex) { status = "수신 중단: " + ex.Message; }
            finally
            {
                foreach (var session in sessions.Values) session.Dispose();
                messages.CompleteAdding();
            }
        }

        private void Pump(Session session, ChatRoomInfo room, DirectRequest watch)
        {
            session.Snapshot.Refresh(cancellation.Token);
            if (session.Snapshot.RefreshDeferred) return;
            if (watch != null && session.Since != watch.Since)
            { session.Cursor = new ChatLogCursor(watch.Since); session.Since = watch.Since; }
            if (!session.OwnChecked)
            {
                using (var db = new ChatSqlite(session.Snapshot.OutputPath))
                    foreach (var row in db.Query("SELECT DISTINCT authorId FROM chatLogs WHERE write_on_pc=1 AND authorId>0 AND type IN (1,26) LIMIT 16")) users.OwnIds.Add(ChatCatalog.Number(row[0]));
                session.OwnChecked = users.OwnIds.Count > 0;
            }
            if (users.OwnIds.Count == 0) throw new IOException("본인 작성자 ID를 확인하지 못했습니다. '나와의 채팅'을 열고 다시 시도하거나 본인 ID를 입력해 주세요.");
            foreach (object[] row in session.Cursor.Read(session.Snapshot.OutputPath))
            {
                foreach (ChatMessage message in ChatMessageDecoder.Decode(room, row, users, watch != null))
                {
                    if (watch != null)
                    {
                        if (message.AuthorId != room.DirectUserId || message.AuthorId != watch.AuthorId) continue;
                    }
                    messages.Add(message, cancellation.Token);
                    if (message.RoomEvent == ChatRoomEvent.Leave || message.RoomEvent == ChatRoomEvent.Kick)
                        ChatEventDiagnostics.Write("received", message);
                }
                session.Cursor.Accept(row);
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
            cancellation.Cancel();
            Completion.ContinueWith(task => { messages.Dispose(); cancellation.Dispose(); }, TaskScheduler.Default);
        }

        private sealed class DirectRequest { public long AuthorId, Since; }
        private sealed class Session : IDisposable
        {
            public KakaoTalkDecryptor.DecryptionKey Key;
            public ChatDatabaseSnapshot Snapshot;
            public ChatLogCursor Cursor = new ChatLogCursor();
            public long Since;
            public bool OwnChecked;
            public Session(KakaoTalkDecryptor.DecryptionKey key, ChatDatabaseSnapshot snapshot) { Key = key; Snapshot = snapshot; }
            public void Dispose() { if (Key != null) Key.Dispose(); }
        }
    }
}
