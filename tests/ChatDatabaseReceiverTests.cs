using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using KakaotalkBot;

internal static class ChatDatabaseReceiverTests
{
    private static int passed;
    private static readonly byte[] RawKey = Enumerable.Range(1, 32).Select(i => (byte)i).ToArray();
    private static readonly byte[] Salt = Enumerable.Range(70, 16).Select(i => (byte)i).ToArray();

    private static int Main(string[] arguments)
    {
        string directory = Path.Combine(Path.GetTempPath(), "ChatReceiverTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            if (arguments.Contains("--live-catalog"))
            {
                var catalog = ChatCatalog.Load(CancellationToken.None);
                foreach (string error in catalog.Errors) Console.WriteLine(error);
                Check(catalog.Rooms.Count > 0, "실제 계정 DB 목록 읽기");
                Console.WriteLine("Live catalog rooms: " + catalog.Rooms.Count);
                VerifyLiveRoom(catalog);
            }
            else Run(directory);
            Console.WriteLine("PASS: " + passed + " receiver checks");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { Directory.Delete(directory, true); }
    }

    private static void VerifyLiveRoom(ChatCatalog catalog)
    {
        var candidates = catalog.Rooms.Where(r => File.Exists(r.SourcePath))
            .OrderByDescending(r => File.GetLastWriteTimeUtc(r.SourcePath)).Take(24).ToArray();
        var decryptor = new KakaoTalkDecryptor();
        foreach (int pid in candidates.Select(r => r.ProcessId).Distinct())
        {
            var keys = decryptor.DiscoverKeys(pid, candidates.Select(r => decryptor.ReadDatabaseInfo(r.SourcePath)));
            try
            {
                var room = candidates.Where(r => keys.ContainsKey(r.SourcePath)).OrderBy(r => new FileInfo(r.SourcePath).Length).FirstOrDefault();
                if (room == null) continue;
                string output = Path.Combine(ChatCatalog.OutputRoot, Path.GetFileName(room.AccountPath), "live-room.sqlite");
                var snapshot = new ChatDatabaseSnapshot(room.SourcePath, output, decryptor, keys[room.SourcePath]);
                var watch = System.Diagnostics.Stopwatch.StartNew();
                snapshot.Refresh(CancellationToken.None);
                Console.WriteLine("Live initial snapshot ms: " + watch.ElapsedMilliseconds);
                using (var db = new ChatSqlite(output)) Check(db.Scalar("SELECT COUNT(*) FROM chatLogs") >= 0, "실제 채팅 테이블 조회");
                var cursor = new ChatLogCursor();
                Check(cursor.Read(output).Count == 0, "실제 DB의 과거 명령 제외");
                watch.Restart();
                snapshot.Refresh(CancellationToken.None);
                Console.WriteLine("Live refresh ms: " + watch.ElapsedMilliseconds + ", decrypted pages: " + snapshot.LastDecryptedPages);
                Check(cursor.Read(output).Count < 256, "실제 DB의 연속 조회");
                using (var receiver = new ChatDatabaseReceiver(room))
                {
                    receiver.Start();
                    var deadline = DateTime.UtcNow.AddSeconds(60);
                    while (!receiver.Status.StartsWith("수신 중") && !receiver.Completion.IsCompleted && DateTime.UtcNow < deadline) Thread.Sleep(100);
                    string status = receiver.Status;
                    receiver.Dispose();
                    Check(receiver.Completion.Wait(TimeSpan.FromSeconds(10)), "실제 수신 작업 종료");
                    Check(status.StartsWith("수신 중"), "실제 수신 서비스 준비: " + status);
                    Console.WriteLine("Live receiver ready and stopped (no command execution)");
                }
                return;
            }
            finally { foreach (var key in keys.Values) key.Dispose(); }
        }
        throw new Exception("실제 채팅방 키가 없어 수신 검증을 완료하지 못했습니다.");
    }

    private static void Run(string directory)
    {
        string plain = Path.Combine(directory, "plain.sqlite");
        string source = Path.Combine(directory, "chatLogs_1.edb");
        string output = Path.Combine(directory, "copy.sqlite");
        using (var db = new ChatSqlite(plain, false))
        {
            int reserve = 80;
            var handle = (IntPtr)typeof(ChatSqlite).GetField("database", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(db);
            Check(sqlite3_file_control(handle, null, 38, ref reserve) == 0, "SQLite 예약 영역 설정");
            db.Query("PRAGMA page_size=4096");
            db.Query("VACUUM");
            db.Query("CREATE TABLE chatLogs(logId INTEGER PRIMARY KEY,authorId INTEGER,type INTEGER,sendAt INTEGER,message TEXT,deleted INTEGER,write_on_pc INTEGER,attachement TEXT)");
            Insert(db, 10, "기존 메시지");
        }
        Check(File.ReadAllBytes(plain)[20] == 80, "실제 SQLite 페이지 형식");
        byte[][] original = EncryptDatabase(plain);
        File.WriteAllBytes(source, original.SelectMany(b => b).ToArray());
        var decryptor = new KakaoTalkDecryptor();
        using (var key = decryptor.CreateKey(decryptor.ReadDatabaseInfo(source), RawKey))
        {
            var snapshot = new ChatDatabaseSnapshot(source, output, decryptor, key);
            Check(snapshot.Refresh(CancellationToken.None), "최초 스냅샷");
            using (var db = new ChatSqlite(output)) Check(db.Scalar("SELECT COUNT(*) FROM chatLogs") == 1, "복호화 사본 SQL 조회");
            Check(!snapshot.Refresh(CancellationToken.None) && snapshot.LastDecryptedPages == 0, "변경 없는 DB 생략");
            var cursor = new ChatLogCursor();
            Check(cursor.Read(output).Count == 0, "최초 시작 시 과거 명령 제외");
            string attachment = "{\"mentions\":[{\"user_id\":9007199254740993,\"at\":[1],\"len\":2}]}";
            using (var db = new ChatSqlite(plain, false)) { Insert(db, 11, "/같은명령", attachment); Insert(db, 12, "/같은명령"); }
            byte[][] second = EncryptDatabase(plain);
            byte[] wal = MakeWal(second, 1);
            File.WriteAllBytes(source + "-wal", wal);
            WriteIndex(source, wal, (uint)second.Length, (uint)second.Length);
            Check(snapshot.Refresh(CancellationToken.None), "WAL 커밋 반영");
            Check(snapshot.LastDecryptedPages > 0 && snapshot.LastDecryptedPages <= second.Length, "변경 페이지만 복호화");
            var rows = cursor.Read(output);
            Check(rows.Count == 2 && (long)rows[0][0] == 11 && (long)rows[1][0] == 12, "같은 내용도 ID별 순서대로 수신");
            Check(Convert.ToString(rows[0][7]) == attachment && rows[1][7] == null, "실제 컬럼명 attachement로 멘션 JSON과 NULL 조회");
            cursor.Accept(11);
            Check(cursor.Read(output).Count == 1, "처리하지 않은 메시지 재조회");
            cursor.Accept(12);
            Check(cursor.Read(output).Count == 0, "중복 제거");
            Check(!snapshot.Refresh(CancellationToken.None), "동일 WAL 중복 반영 방지");

            // 미커밋 프레임이 존재해도 공유 인덱스의 마지막 커밋까지만 사용합니다.
            using (var db = new ChatSqlite(plain, false)) Insert(db, 13, "다음 명령");
            byte[][] third = EncryptDatabase(plain);
            byte[] extended = AppendWal(wal, third, true);
            File.WriteAllBytes(source + "-wal", extended);
            Check(!snapshot.Refresh(CancellationToken.None), "공개되지 않은 WAL 트랜잭션 제외");
            WriteIndex(source, extended, (uint)(second.Length + third.Length), (uint)third.Length);
            Check(snapshot.Refresh(CancellationToken.None), "새 커밋 이후 프레임만 갱신");
            Check(cursor.Read(output).Count == 1 && (long)cursor.Read(output)[0][0] == 13, "증분 조회");
            cursor.Accept(13);

            // 체크포인트 후 WAL 세대가 바뀌면 원본을 대조하되 같은 페이지의 복호화는 생략합니다.
            File.WriteAllBytes(source, third.SelectMany(b => b).ToArray());
            byte[] reset = MakeWal(new byte[0][], 2);
            File.WriteAllBytes(source + "-wal", reset);
            WriteIndex(source, reset, 0, (uint)third.Length);
            Check(snapshot.Refresh(CancellationToken.None) && snapshot.LastDecryptedPages == 0, "체크포인트/세대 변경과 키 재사용");
            Check(cursor.Read(output).Count == 0, "체크포인트 후 명령 재실행 방지");

            using (var db = new ChatSqlite(plain, false)) Insert(db, 14, "오류 후 재시도");
            byte[][] fourth = EncryptDatabase(plain);
            byte[] next = MakeWal(fourth, 2);
            byte[] corrupted = (byte[])next.Clone();
            corrupted[60] ^= 1;
            File.WriteAllBytes(source + "-wal", corrupted);
            WriteIndex(source, next, (uint)fourth.Length, (uint)fourth.Length);
            Expect<IOException>(() => snapshot.Refresh(CancellationToken.None), "WAL 체크섬 오류 거부");
            using (var db = new ChatSqlite(output)) Check(db.Scalar("SELECT MAX(logId) FROM chatLogs") == 13, "실패 시 이전 스냅샷 보존");
            File.WriteAllBytes(source + "-wal", next);
            Check(snapshot.Refresh(CancellationToken.None), "복구 후 재시도");
            Check(cursor.Read(output).Count == 1, "실패 중 수신 기준점 유지");
            cursor.Accept(14);

            using (var shm = new FileStream(source + "-shm", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
            {
                var overlap = new Overlap { Offset = 120 };
                Check(LockFileEx(shm.SafeFileHandle, 3, 0, 3, 0, ref overlap), "테스트 쓰기/체크포인트/복구 잠금");
                try { Check(!snapshot.Refresh(CancellationToken.None), "카카오톡 잠금과 충돌 없이 기존 커밋 확인"); }
                finally { UnlockFileEx(shm.SafeFileHandle, 0, 3, 0, ref overlap); }
            }
            string[] originals = { source, source + "-wal", source + "-shm" };
            byte[][] originalBytes = originals.Select(File.ReadAllBytes).ToArray();
            DateTime[] originalTimes = originals.Select(File.GetLastWriteTimeUtc).ToArray();
            var independent = new ChatDatabaseSnapshot(source, output + ".check", decryptor, key);
            Check(independent.Refresh(CancellationToken.None), "원본 수정 없이 전체 사본 생성");
            for (int i = 0; i < originals.Length; i++)
                Check(originalBytes[i].SequenceEqual(File.ReadAllBytes(originals[i])) &&
                    originalTimes[i] == File.GetLastWriteTimeUtc(originals[i]), "원본 DB/WAL/SHM 내용과 수정 시각 보존");

            Type readerType = typeof(ChatDatabaseSnapshot).GetNestedType("SourceReader", BindingFlags.NonPublic);
            using (var reader = (IDisposable)Activator.CreateInstance(readerType, new object[] { source }))
            {
                using (var shm = new FileStream(source + "-shm", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
                {
                    var overlap = new Overlap { Offset = 120 };
                    Check(LockFileEx(shm.SafeFileHandle, 3, 0, 3, 0, ref overlap), "봇이 읽는 동안 카카오톡의 쓰기 잠금 획득 가능");
                    UnlockFileEx(shm.SafeFileHandle, 0, 3, 0, ref overlap);
                    shm.Position = 128;
                    shm.Write(BitConverter.GetBytes(1U), 0, 4);
                }
                bool rejected = !(bool)readerType.GetMethod("Verify").Invoke(reader, null);
                Check(rejected, "읽는 중 체크포인트 변경 거부");
            }
            File.WriteAllBytes(source + "-shm", originalBytes[2]);
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.Cancel();
                Expect<OperationCanceledException>(() => snapshot.Refresh(cancellation.Token), "수신 취소");
            }
            Check(!snapshot.Refresh(CancellationToken.None), "취소 후 재시도");
        }

        string usersPath = Path.Combine(directory, "users.sqlite"), roomsPath = Path.Combine(directory, "rooms.sqlite");
        using (var db = new ChatSqlite(usersPath, false))
        {
            db.Query("CREATE TABLE talkUser(userId INTEGER,linkId INTEGER,nickName TEXT,friendNickName TEXT)");
            db.Query("INSERT INTO talkUser VALUES(7,99,'오픈이름','친구이름'),(7,0,'일반이름','친구이름'),(8,99,'다른사람','')");
        }
        using (var db = new ChatSqlite(roomsPath, false))
        {
            db.Query("CREATE TABLE chatRoomList(chatId INTEGER,type TEXT)");
            db.Query("CREATE TABLE chatMembers(chatId INTEGER,userId INTEGER)");
            db.Query("INSERT INTO chatRoomList VALUES(1,'MemoChat')");
            db.Query("INSERT INTO chatMembers VALUES(1,9)");
        }
        var users = new ChatUserDirectory();
        users.Load(usersPath, roomsPath);
        Check(users.Find(7, 99) == "오픈이름" && users.Find(7, 0) == "친구이름", "방 프로필별 닉네임");
        Check(users.Find(7, 123) == null, "잘못된 프로필로 대체하지 않음");
        var room = new ChatRoomInfo { ChatId = 50, LinkId = 99 };
        CheckMentions(room, users);
        object[] text = { 20L, 7L, 1L, 100L, "/명령\n다음 줄", 0L, 0L };
        Check(ChatMessageDecoder.Decode(room, text, users, false).Single().Message.Contains("\n"), "여러 줄 메시지 유지");
        text[2] = 26L;
        Check(ChatMessageDecoder.Decode(room, text, users, false).Count == 1, "답글 명령");
        text[2] = 16385L;
        Check(ChatMessageDecoder.Decode(room, text, users, false).Count == 0, "삭제된 메시지 타입 제외");
        text[2] = 1L; text[1] = 9L;
        Check(ChatMessageDecoder.Decode(room, text, users, false).Single().IsOwn, "자기 메시지는 계정 등록용으로 전달");
        text[1] = 7L; text[5] = 1L;
        Check(ChatMessageDecoder.Decode(room, text, users, false).Count == 0, "삭제 표시 제외");
        text[5] = 0L; text[1] = 999L;
        Check(ChatMessageDecoder.Decode(room, text, users, false).Single().AuthorId == 999, "프로필 미확인 사용자도 ID로 수신하고 후속 이벤트를 막지 않음");
        text[1] = 0L; text[2] = 0L;
        text[4] = "{\"feedType\":4,\"members\":[{\"userId\":7,\"nickName\":\"입장자\"}]}";
        Check(ChatMessageDecoder.Decode(room, text, users, false).Single().EventCommand == "/입장", "입장 이벤트");
        text[4] = "{\"feedType\":2,\"member\":{\"userId\":7,\"nickName\":\"퇴장자\"}}";
        Check(ChatMessageDecoder.Decode(room, text, users, false).Single().EventCommand == "/퇴장", "퇴장 이벤트");
        Check(ChatMessageDecoder.Decode(room, text, users, true).Count == 0, "1:1 이벤트 명령 실행 방지");
        CheckSystemEvents(room, users);
        CheckDelayedSystemEvents(directory, room, users);
        string rosterPath = Path.Combine(directory, "roster.sqlite");
        using (var db = new ChatSqlite(rosterPath, false))
        {
            db.Query("CREATE TABLE chatRoomList(chatId INTEGER,lastLogId INTEGER)");
            db.Query("CREATE TABLE chatMembers(chatId INTEGER,userId INTEGER,isActive INTEGER)");
            db.Query("INSERT INTO chatRoomList VALUES(50,123)");
            db.Query("INSERT INTO chatMembers VALUES(50,7,1),(50,8,0),(50,0,1),(50,10,NULL),(999,11,1)");
        }
        var roster = ChatRosterSnapshot.Read(rosterPath, room, users);
        Check(roster.ChatId == 50 && roster.LogId == 123 && roster.Members.Count == 2, "방별 참여자 조회와 무효/미확인 값 제외");
        Check(roster.Members.Single(m => m.UserId == 7).IsPresent && !roster.Members.Single(m => m.UserId == 8).IsPresent,
            "isActive로 참여 중·퇴장 상태 구분");
        Check(roster.Members.All(m => m.MemberType == -1), "역할 컬럼이 없으면 권한 미확인 처리");
        using (var db = new ChatSqlite(usersPath, false))
        {
            db.Query("ALTER TABLE talkUser ADD COLUMN openMemberType INTEGER");
            db.Query("ALTER TABLE talkUser ADD COLUMN openMemberPrivilege TEXT");
            db.Query("UPDATE talkUser SET openMemberType=4,openMemberPrivilege='18446744073709551615' WHERE userId=7 AND linkId=99");
            db.Query("UPDATE talkUser SET openMemberType=1 WHERE userId=7 AND linkId=0");
        }
        users.Load(usersPath, roomsPath);
        var roles = ChatRosterSnapshot.Read(rosterPath, room, users);
        Check(roles.Members.Single(m => m.UserId == 7).MemberType == 4, "프로필 DB 역할이 참여자 스냅샷으로 전달");
        Check(roles.Members.Single(m => m.UserId == 7).MemberPrivilege == "18446744073709551615", "원본 권한 값 문자열 보존");
        Check(users.MemberType(7, 0) == -1 && users.MemberType(7, 123) == -1, "다른 방 역할 사용 금지");

        var dm = new ChatLogCursor(100);
        var dmRows = dm.Read(output);
        Check(dmRows.Count == 5, "1:1 요청 시각 이후 메시지 조회");
    }

    private static void CheckDelayedSystemEvents(string directory, ChatRoomInfo room, ChatUserDirectory users)
    {
        string path = Path.Combine(directory, "delayed-events.sqlite");
        using (var db = new ChatSqlite(path, false))
        {
            db.Query("CREATE TABLE chatLogs(logId INTEGER PRIMARY KEY,authorId INTEGER,type INTEGER,sendAt INTEGER,message TEXT,deleted INTEGER,write_on_pc INTEGER,attachement TEXT)");
            string kick = "{\"feedType\":6,\"member\":{\"userId\":7160322422266675266,\"nickName\":\"케이\"}}";
            db.Query("INSERT INTO chatLogs VALUES(1,9,0,100,?,0,0,NULL)", kick);
            var cursor = new ChatLogCursor();
            Check(cursor.Read(path).Count == 0, "시작 이전 강퇴는 소급 집계하지 않음");
            db.Query("INSERT INTO chatLogs VALUES(3,9,0,100,'{}',0,0,NULL),(4,7,1,100,'일반 메시지',0,0,NULL)");
            var first = cursor.Read(path);
            foreach (var row in first) cursor.Accept(row);
            db.Query("UPDATE chatLogs SET message=? WHERE logId=3", kick);
            db.Query("INSERT INTO chatLogs VALUES(2,9,0,100,?,0,0,NULL),(5,7,1,100,'다음 메시지',0,0,NULL)", kick);
            var next = cursor.Read(path);
            Check(next.Select(row => ChatCatalog.Number(row[0])).SequenceEqual(new long[] { 2, 3, 5 }), "새 채팅이 있어도 지연 삽입·본문 갱신된 강퇴를 다시 읽음");
            var events = next.SelectMany(row => ChatMessageDecoder.Decode(room, row, users, false)).Where(message => message.RoomEvent == ChatRoomEvent.Kick).ToArray();
            Check(events.Length == 2 && events.All(message => message.AuthorId == 7160322422266675266L), "프로필에 없는 대상의 지연 강퇴 해석");
            foreach (var row in next) cursor.Accept(row);
            Check(cursor.Read(path).Count == 0, "재확인한 시스템 행을 중복 전달하지 않음");
        }
    }

    private static void CheckSystemEvents(ChatRoomInfo room, ChatUserDirectory users)
    {
        object[] row = { 30L, 99L, 0L, 100L, "{\"feedType\":2,\"member\":{\"userId\":7,\"nickName\":\"동일닉\"}}", 0L, 0L };
        var leave = ChatMessageDecoder.Decode(room, row, users, false).Single();
        Check(leave.RoomEvent == ChatRoomEvent.Leave && leave.AuthorId == 7, "퇴장 대상 ID는 시스템 작성자가 아닌 member에서 읽음");
        row[1] = 9L;
        row[4] = "{\"feedType\":6,\"member\":{\"userId\":8,\"nickName\":\"동일닉\"}}";
        var kick = ChatMessageDecoder.Decode(room, row, users, false).Single();
        Check(kick.RoomEvent == ChatRoomEvent.Kick && kick.AuthorId == 8 && kick.EventCommand == "/퇴장", "본인이 내보낸 대상도 강퇴로 구분하고 기존 퇴장 응답 유지");
        row[4] = "{\"feedType\":6,\"member\":{\"userId\":\"9007199254740993\"}}";
        kick = ChatMessageDecoder.Decode(room, row, users, false).Single();
        Check(kick.AuthorId == 9007199254740993L && kick.Nickname == null, "이름 없는 대상과 큰 ID도 집계 전달");
        row[4] = "{\"feedType\":2,\"members\":[{\"userId\":7},{\"userId\":8},{\"userId\":-1},{\"userId\":\"1.2\"}],\"member\":{\"userId\":7}}";
        var members = ChatMessageDecoder.Decode(room, row, users, false);
        Check(members.Count == 2 && members.All(message => message.RoomEvent == ChatRoomEvent.Leave), "한 이벤트의 중복 대상과 잘못된 ID 제외");
        row[4] = "{\"feedType\":6,\"hidden\":true,\"member\":{\"userId\":8}}";
        Check(ChatMessageDecoder.Decode(room, row, users, false).Count == 0, "숨겨진 이벤트 제외");
        row[4] = "{\"feedType\":11,\"member\":{\"userId\":8}}";
        var granted = ChatMessageDecoder.Decode(room, row, users, false).Single();
        Check(granted.RoleChanged && granted.ExpectedMemberType == 4 && granted.AuthorId == 8 && granted.RoomEvent == ChatRoomEvent.None,
            "관리자 지정은 권한 갱신으로 전달하고 강퇴로 집계하지 않음");
        row[1] = 7L; row[2] = 1L; row[4] = "동일닉님을 내보냈습니다.";
        Check(ChatMessageDecoder.Decode(room, row, users, false).Single().RoomEvent == ChatRoomEvent.None, "일반 채팅의 시스템 문구는 집계하지 않음");
    }

    private static void CheckMentions(ChatRoomInfo room, ChatUserDirectory users)
    {
        object[] row = { 21L, 7L, 1L, 100L, "/명령 @가 @나 @가", 0L, 0L, null };
        Func<ChatMessage> decode = () => ChatMessageDecoder.Decode(room, row, users, false).Single();
        Check(decode().Mentions.Count == 0, "첨부 정보 없는 메시지의 빈 멘션 목록");
        row[7] = "{\"mentions\":[{\"user_id\":8,\"at\":[2],\"len\":1},{\"user_id\":9007199254740993,\"at\":[3,1,1],\"len\":1}]}";
        var message = decode();
        Check(message.Mentions.Count == 2 && message.Mentions[0].UserId == 9007199254740993L, "큰 64비트 사용자 ID 보존 및 등장순 정렬");
        Check(message.Mentions[0].At.SequenceEqual(new[] { 1, 3 }) && message.Mentions[0].Length == 1, "반복 멘션 순번과 표시 길이 보존");
        Check(message.Mentions[0].Nickname == null && message.Mentions[1].Nickname == "다른사람", "프로필 없는 대상도 ID 전달 및 방 프로필 이름 조회");
        row[7] = "{\"mentions\":[{\"user_id\":8,\"at\":[1],\"len\":1},{\"user_id\":\"8\",\"at\":[3],\"len\":2},{\"user_id\":\"9007199254740993\",\"at\":[2],\"len\":1}]}";
        var command = new Command { Mentions = decode().Mentions };
        Check(command.Mentions.Count == 3 && command.MentionedUserIds.SequenceEqual(new long[] { 8, 9007199254740993L }), "숫자 문자열 ID 처리 및 명령 대상 ID 중복 제거");
        row[2] = 26L;
        row[7] = "{\"src_mentions\":[{\"user_id\":999,\"at\":[1],\"len\":2}],\"mentions\":[{\"user_id\":8,\"at\":[1],\"len\":1}]}";
        Check(decode().Mentions.Single().UserId == 8, "답장에서는 현재 메시지의 멘션만 전달");
        row[7] = "{\"src_mentions\":[{\"user_id\":999,\"at\":[1],\"len\":2}]}";
        Check(decode().Mentions.Count == 0, "인용 원본의 멘션을 현재 태그로 오인하지 않음");
        foreach (string invalid in new[] { "", "{", "[]", "null", "{\"mentions\":\"잘못된 형식\"}" })
        {
            row[7] = invalid;
            Check(decode().Message == Convert.ToString(row[4]) && decode().Mentions.Count == 0, "잘못된 메타데이터도 본문 수신 유지");
        }
        row[7] = "{\"mentions\":[null,1,{\"user_id\":9223372036854775808,\"at\":[1],\"len\":1},{\"user_id\":1.5,\"at\":[1],\"len\":1},{\"user_id\":8,\"at\":[0,-1],\"len\":1},{\"user_id\":8,\"at\":[1],\"len\":-1},{\"user_id\":8,\"at\":[1],\"len\":1}]}";
        Check(decode().Mentions.Count == 1 && decode().Mentions[0].UserId == 8, "유효하지 않은 항목만 제외하고 정상 멘션 유지");
        Check(new Command().Mentions.Count == 0 && new Command().MentionedUserIds.Count == 0, "기본 명령의 멘션 API는 null 대신 빈 목록 반환");
    }

    private static void Insert(ChatSqlite db, long id, string message, string attachment = null)
    { db.Query("INSERT INTO chatLogs VALUES(?,7,1,100,?,0,0,?)", id, message, attachment); }

    private static byte[][] EncryptDatabase(string path)
    {
        byte[] plain = File.ReadAllBytes(path);
        return Enumerable.Range(0, plain.Length / 4096).Select(i => EncryptPage(plain.Skip(i * 4096).Take(4096).ToArray(), i + 1)).ToArray();
    }

    private static byte[] EncryptPage(byte[] plain, int number)
    {
        byte[] page = new byte[4096];
        int start = number == 1 ? 16 : 0;
        if (start != 0) Buffer.BlockCopy(Salt, 0, page, 0, 16);
        byte[] iv = Enumerable.Repeat((byte)43, 16).ToArray();
        using (var aes = Aes.Create())
        {
            aes.Key = RawKey; aes.IV = iv; aes.Mode = CipherMode.CBC; aes.Padding = PaddingMode.None;
            using (var encrypt = aes.CreateEncryptor()) Buffer.BlockCopy(encrypt.TransformFinalBlock(plain, start, 4016 - start), 0, page, start, 4016 - start);
        }
        Buffer.BlockCopy(iv, 0, page, 4016, 16);
        using (var pbkdf = new Rfc2898DeriveBytes(RawKey, Salt.Select(b => (byte)(b ^ 0x3a)).ToArray(), 2, HashAlgorithmName.SHA512))
        using (var hmac = new HMACSHA512(pbkdf.GetBytes(32)))
            Buffer.BlockCopy(hmac.ComputeHash(page.Skip(start).Take(4032 - start).Concat(BitConverter.GetBytes(number)).ToArray()), 0, page, 4032, 64);
        return page;
    }

    private static byte[] MakeWal(byte[][] pages, uint generation)
    {
        byte[] header = new byte[32];
        Big(header, 0, 0x377f0682); Big(header, 4, 3007000); Big(header, 8, 4096);
        Big(header, 16, generation); Big(header, 20, 345);
        uint s1 = 0, s2 = 0;
        Sum(header, 0, 24, ref s1, ref s2); Big(header, 24, s1); Big(header, 28, s2);
        return AppendWal(header, pages, true);
    }

    private static byte[] AppendWal(byte[] wal, byte[][] pages, bool commit)
    {
        uint s1 = ChatDatabaseSnapshot.Big(wal, wal.Length == 32 ? 24 : wal.Length - 4096 - 8);
        uint s2 = ChatDatabaseSnapshot.Big(wal, wal.Length == 32 ? 28 : wal.Length - 4096 - 4);
        using (var stream = new MemoryStream())
        {
            stream.Write(wal, 0, wal.Length);
            for (int i = 0; i < pages.Length; i++)
            {
                byte[] header = new byte[24];
                Big(header, 0, (uint)i + 1); Big(header, 4, commit && i == pages.Length - 1 ? (uint)pages.Length : 0);
                Buffer.BlockCopy(wal, 16, header, 8, 8);
                Sum(header, 0, 8, ref s1, ref s2); Sum(pages[i], 0, 4096, ref s1, ref s2);
                Big(header, 16, s1); Big(header, 20, s2);
                stream.Write(header, 0, 24); stream.Write(pages[i], 0, 4096);
            }
            return stream.ToArray();
        }
    }

    private static void WriteIndex(string source, byte[] wal, uint frames, uint pages)
    {
        byte[] header = new byte[48];
        Buffer.BlockCopy(BitConverter.GetBytes(3007000U), 0, header, 0, 4);
        header[12] = 1; header[15] = 16;
        Buffer.BlockCopy(BitConverter.GetBytes(frames), 0, header, 16, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(pages), 0, header, 20, 4);
        Buffer.BlockCopy(wal, 16, header, 32, 8);
        uint s1 = 0, s2 = 0; Sum(header, 0, 40, ref s1, ref s2);
        Buffer.BlockCopy(BitConverter.GetBytes(s1), 0, header, 40, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(s2), 0, header, 44, 4);
        File.WriteAllBytes(source + "-shm", header.Concat(header).Concat(new byte[40]).ToArray());
    }

    private static void Sum(byte[] bytes, int offset, int count, ref uint a, ref uint b)
    {
        unchecked { for (int i = offset; i < offset + count; i += 8) { a += BitConverter.ToUInt32(bytes, i) + b; b += BitConverter.ToUInt32(bytes, i + 4) + a; } }
    }
    private static void Big(byte[] bytes, int offset, uint value) { for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (24 - 8 * i)); }
    private static void Check(bool condition, string message) { if (!condition) throw new Exception("FAIL: " + message); passed++; }
    private static void Expect<T>(Action action, string message) where T : Exception
    { try { action(); } catch (T) { passed++; return; } throw new Exception("FAIL: " + message); }

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern int sqlite3_file_control(IntPtr database, byte[] name, int operation, ref int value);
    [StructLayout(LayoutKind.Sequential)] private struct Overlap { public IntPtr A, B; public uint Offset, OffsetHigh; public IntPtr Event; }
    [DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool LockFileEx(Microsoft.Win32.SafeHandles.SafeFileHandle file, uint flags, uint reserved, uint low, uint high, ref Overlap overlap);
    [DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnlockFileEx(Microsoft.Win32.SafeHandles.SafeFileHandle file, uint reserved, uint low, uint high, ref Overlap overlap);
}
