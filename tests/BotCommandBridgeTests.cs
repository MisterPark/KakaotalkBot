using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using KakaotalkBot;

internal static class BotCommandBridgeTests
{
    private static int Main()
    {
        try
        {
            // 생성자의 보이스룸 이미지 읽기와 외부 서비스 초기화를 제외하고 명령 연결만 검증합니다.
            var bot = (Bot)FormatterServices.GetUninitializedObject(typeof(Bot));
            var commands = new Queue<Command>();
            var answers = new Queue<Bot.QuizAnswer>();
            Set(bot, "commands", commands);
            Set(bot, "quizAnswers", answers);
            Set(bot, "chatLog", new List<string>());
            var room = new ChatRoomInfo();
            typeof(ChatRoomInfo).GetProperty("ChatId").SetValue(room, 10L, null);
            typeof(Bot).GetProperty("SelectedRoom").SetValue(bot, room, null);
            Database.Instance.Keywords.Add("/테스트");
            Database.Instance.AddUser(7, "검증사용자");
            Type messageType = typeof(Bot).Assembly.GetType("KakaotalkBot.ChatMessage");
            MethodInfo handle = typeof(Bot).GetMethod("HandleIncomingMessage", BindingFlags.Instance | BindingFlags.NonPublic);
            var mentions = new[]
            {
                new ChatMention(9007199254740993L, new[] { 1, 3 }, 3, "동일닉"),
                new ChatMention(8, new[] { 2 }, 3, "동일닉")
            };
            Action<long, long, string, string> deliver = (chat, log, text, eventCommand) =>
            {
                object message = Activator.CreateInstance(messageType, true);
                messageType.GetField("ChatId").SetValue(message, chat);
                messageType.GetField("LogId").SetValue(message, log);
                messageType.GetField("AuthorId").SetValue(message, 7L);
                messageType.GetField("Nickname").SetValue(message, "검증사용자");
                messageType.GetField("Message").SetValue(message, text);
                messageType.GetField("EventCommand").SetValue(message, eventCommand);
                if (log == 100 || eventCommand != null) messageType.GetField("Mentions").SetValue(message, mentions);
                handle.Invoke(bot, new[] { message });
            };
            deliver(10, 100, "/테스트", null);
            deliver(10, 101, "/테스트", null);
            deliver(10, 102, "퀴즈 답변", null);
            deliver(999, 103, "/테스트", null);
            deliver(10, 104, "/입장", "/입장");
            if (commands.Count != 3 || answers.Count != 3) throw new Exception("명령/퀴즈 분기 또는 방 격리 오류");
            if (answers.Any(answer => answer.AuthorId != 7)) throw new Exception("퀴즈 작성자 ID 누락");
            var values = commands.ToArray();
            if (values[0].LogId != 100 || values[1].LogId != 101 || values[2].Keyword != "/입장") throw new Exception("명령 순서 오류");
            if (values.Any(c => c.AuthorId != 7 || c.ChatId != 10 || c.Nickname != "검증사용자")) throw new Exception("명령 작성자 정보 오류");
            if (Database.Instance.UserTable[0].Contribution != 3) throw new Exception("기여도 중복/누락 오류");
            if (!values[0].MentionedUserIds.SequenceEqual(new long[] { 9007199254740993L, 8 })) throw new Exception("멘션 대상 ID 전달 오류");
            if (values[0].Mentions.Any(m => m.Nickname != "동일닉") || !values[0].Mentions[0].At.SequenceEqual(new[] { 1, 3 })) throw new Exception("동일 닉네임의 사용자 ID 구분 또는 반복 멘션 정보 오류");
            if (values[1].Mentions.Count != 0 || values[2].Mentions.Count != 0) throw new Exception("일반 메시지 또는 입퇴장 이벤트로 멘션 정보가 유출됨");
            var copy = values[0];
            var mutableMentions = new List<ChatMention>(mentions);
            copy.Mentions = mutableMentions;
            mutableMentions.Clear();
            if (copy.Mentions.Count != 2) throw new Exception("명령이 외부 멘션 목록 변경에 영향받음");
            deliver(10, 105, "/메모 @대상 사유", null);
            if (commands.Count != 4 || commands.Last().Keyword != "/메모 @대상 사유" || commands.Last().ReceivedAt <= 0)
                throw new Exception("운영진 명령이 키워드 시트 등록 없이 수신되지 않음");
            Set(bot, "isBotRunning", true);
            int queuedBeforeFailure = commands.Count;
            typeof(Bot).GetMethod("DeferProcessing", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(bot, new object[] { new System.IO.IOException("시험 저장 실패") });
            if (!bot.IsBotRunning || commands.Count != queuedBeforeFailure || bot.SelectedRoom != room)
                throw new Exception("복구 가능한 오류가 수신 상태나 명령 큐를 초기화함");
            bot.Update();
            if (!bot.IsBotRunning || commands.Count != queuedBeforeFailure || bot.LastProcessingError == null)
                throw new Exception("재시도 대기 중 수신 상태·큐 보존 실패");
            Set(bot, "isBotRunning", false);
            long privateId = 9007199254740993L;
            Database.Instance.AddUser(privateId, "검증대상");
            var filter = typeof(Database).GetMethod("ForChat", BindingFlags.Instance | BindingFlags.NonPublic);
            string publicText = (string)filter.Invoke(Database.Instance, new object[] {
                "사용자 ID: " + privateId + "\n요청자 ID: " + privateId + "\n사용자 " + privateId +
                " · ID " + privateId + "\n작성자: " + privateId + "\n메모 " + privateId + "\n포인트 100 / 레벨 2" });
            if (publicText.Contains(privateId.ToString()) || !publicText.Contains("포인트 100 / 레벨 2"))
                throw new Exception("채팅 ID 비공개 처리 오류");
            string internalHistory = Database.Instance.OperatorHistory(10, privateId);
            string publicHistory = Database.Instance.OperatorHistory(10, privateId, false);
            if (!internalHistory.Contains(privateId.ToString()) || publicHistory.Contains(privateId.ToString()))
                throw new Exception("내부 이력/채팅 이력 ID 표시 분리 오류");
            object activity = typeof(Database).GetField("activity", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Database.Instance);
            var observe = activity.GetType().GetMethod("Observe", BindingFlags.Instance | BindingFlags.NonPublic);
            observe.Invoke(activity, new object[] { 10L, privateId, "첫방이름", true, 1900000000L, 1000L, "chat" });
            observe.Invoke(activity, new object[] { 20L, privateId, "둘째방이름", true, 1900000001L, 1001L, "chat" });
            if (Database.Instance.ChatUserName(10, privateId) != "첫방이름" || Database.Instance.ChatUserName(20, privateId) != "둘째방이름")
                throw new Exception("방별 닉네임 혼합");
            observe.Invoke(activity, new object[] { 10L, privateId, "첫방변경", true, 1900000002L, 1002L, "chat" });
            if (Database.Instance.ChatUserName(20, privateId) != "둘째방이름" ||
                Database.Instance.NicknameHistory.Count(n => n.UserId == privateId && n.ChatId == 10) != 1 ||
                Database.Instance.NicknameHistory.Any(n => n.UserId == privateId && n.ChatId == 20))
                throw new Exception("다른 방 이름으로 변경 이력 생성");
            if (Database.Instance.ChatUserName(30, privateId) != "이름 미확인 사용자")
                throw new Exception("알 수 없는 방에서 전역 닉네임 노출");
            Database.Instance.GetOrAddUser(privateId).Experience = 200;
            Database.Instance.AddUser(8101, "다른방전역이름").Experience = 200;
            observe.Invoke(activity, new object[] { 10L, 8101L, "공동순위", true, 1900000003L, 1003L, "chat" });
            Database.Instance.AddUser(8102, "퇴장사용자").Experience = 9999;
            observe.Invoke(activity, new object[] { 10L, 8102L, "퇴장사용자", false, 1900000003L, 1003L, "chat" });
            string levelRanking = Database.Instance.LevelRanking(10);
            if (!levelRanking.Contains("1위 첫방변경") || !levelRanking.Contains("1위 공동순위") || levelRanking.Contains("퇴장사용자") || levelRanking.Contains("다른방전역이름"))
                throw new Exception("레벨 공동 순위·방별 닉네임·퇴장 제외 오류");
            object operations = typeof(Database).GetField("Operations", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Database.Instance);
            var recordType = typeof(Bot).Assembly.GetType("KakaotalkBot.OperationRecord");
            Action<long, long, long, string> addMonthly = (chatId, idValue, count, month) => {
                object record = Activator.CreateInstance(recordType, true);
                foreach (var pair in new Dictionary<string, object> { { "Kind", "monthly" }, { "Key", "rank-test:" + chatId + ":" + idValue + ":" + month },
                    { "ChatId", chatId }, { "UserId", idValue }, { "Count", count }, { "Month", month } }) recordType.GetField(pair.Key).SetValue(record, pair.Value);
                operations.GetType().GetMethod("Put", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(operations, new[] { record });
            };
            addMonthly(10, privateId, 30, "2026-09"); addMonthly(10, 8101, 30, "2026-09");
            addMonthly(20, 8102, 999, "2026-09"); addMonthly(10, 8102, 999, "2026-08");
            string personal = Database.Instance.PersonalRankings(10, privateId, "2026-09");
            if (!personal.Contains("1위 / 2명 (30회)") || !personal.Contains("레벨 순위: 1위")) throw new Exception("월·방 분리 개인 순위 오류");
            if (!Database.Instance.PersonalRankings(10, 8102, "2026-09").Contains("순위 없음 (0회)")) throw new Exception("채팅 없는 사용자 순위 오류");
            deliver(10, 110, "/레벨랭킹", null);
            if (commands.Last().Keyword != "/레벨랭킹") throw new Exception("레벨랭킹 명령 수신 오류");
            int queuedBeforeOwn = commands.Count, answersBeforeOwn = answers.Count;
            object ownMessage = Activator.CreateInstance(messageType, true);
            foreach (var pair in new Dictionary<string, object> { { "ChatId", 10L }, { "LogId", 2000L }, { "AuthorId", 8200L },
                { "SendAt", 1900000010L }, { "Nickname", "봇계정" }, { "Message", "/테스트" }, { "IsOwn", true } })
                messageType.GetField(pair.Key).SetValue(ownMessage, pair.Value);
            handle.Invoke(bot, new[] { ownMessage });
            User ownUser;
            if (!Database.Instance.FindUser(8200, out ownUser) || Database.Instance.ChatUserName(10, 8200) != "봇계정")
                throw new Exception("봇 계정 등록 또는 방별 닉네임 누락");
            messageType.GetField("Message").SetValue(ownMessage, "자동 응답");
            messageType.GetField("LogId").SetValue(ownMessage, 2001L);
            handle.Invoke(bot, new[] { ownMessage });
            if (commands.Count != queuedBeforeOwn || answers.Count != answersBeforeOwn || ownUser.Experience != 0 || ownUser.Contribution != 0 ||
                !Database.Instance.PersonalRankings(10, 8200, "2030-03").Contains("순위 없음 (0회)"))
                throw new Exception("봇 출력이 명령·퀴즈·보상으로 재처리됨");
            // 실제 창 조작 없이 주기, 초안 보호, 실패 후 재시도와 수신 상태 유지를 검증합니다.
            var maintain = typeof(Bot).GetMethod("MaintainChatWindow", BindingFlags.Instance | BindingFlags.NonPublic);
            DateTime now = new DateTime(2026, 9, 24, 0, 0, 0, DateTimeKind.Utc);
            Set(bot, "isBotRunning", true);
            Set(bot, "targetWindow", "검증방");
            Set(bot, "nextRoomRecycle", now.AddHours(4));
            bool previousAuto = StaticVariable.AutoReboot;
            StaticVariable.AutoReboot = true;
            int recycled = 0;
            Func<bool> recycle = () => { recycled++; return true; };
            maintain.Invoke(bot, new object[] { now.AddHours(5), recycle });
            if (recycled != 0 || commands.Count != queuedBeforeOwn) throw new Exception("명령 대기 중 재열기 또는 큐 손실");
            commands.Clear();
            maintain.Invoke(bot, new object[] { now.AddHours(3), recycle });
            if (recycled != 0) throw new Exception("4시간 이전 재열기");
            StaticVariable.AutoReboot = false;
            maintain.Invoke(bot, new object[] { now.AddHours(4), recycle });
            if (recycled != 0) throw new Exception("자동 재열기 해제 무시");
            StaticVariable.AutoReboot = true;
            maintain.Invoke(bot, new object[] { now.AddHours(4), new Func<bool>(() => false) });
            if (bot.RoomRecycleStatus == null || bot.LastRoomRecycle != DateTime.MinValue) throw new Exception("초안 대기 처리 실패");
            maintain.Invoke(bot, new object[] { now.AddHours(4).AddSeconds(30), recycle });
            if (recycled != 0) throw new Exception("초안 대기 재시도 간격 오류");
            maintain.Invoke(bot, new object[] { now.AddHours(4).AddMinutes(1), new Func<bool>(() => { throw new Exception("시험"); }) });
            if (!bot.IsBotRunning || !bot.RoomRecycleStatus.Contains("시험")) throw new Exception("재열기 실패 시 수신 유지 실패");
            maintain.Invoke(bot, new object[] { now.AddHours(4).AddMinutes(2), recycle });
            if (recycled != 1 || !bot.IsBotRunning || bot.RoomRecycleStatus != null || answers.Count != answersBeforeOwn)
                throw new Exception("재열기 성공 후 상태 유지 실패");
            maintain.Invoke(bot, new object[] { now.AddHours(8).AddMinutes(1), recycle });
            if (recycled != 1) throw new Exception("성공 후 4시간 주기 오류");
            deliver(10, 3001, "/퀘스트", null);
            deliver(10, 3002, "/일퀘", null);
            deliver(10, 3003, "/출석", null);
            deliver(10, 3004, "/출석체크", null);
            deliver(10, 3005, "/퀴즈", null);
            deliver(10, 3006, "/과학퀴즈", null);
            deliver(10, 3007, "/경제퀴즈", null);
            deliver(10, 3008, "/퀴즈목록", null);
            deliver(10, 3009, "/퀴즈등록", null);
            deliver(10, 3010, "/퀴즈삭제 과학 | 과학 문제", null);
            if (!commands.Select(c => c.Keyword).SequenceEqual(new[] { "/퀘스트", "/일퀘", "/출첵", "/출첵", "/상식퀴즈", "/과학퀴즈", "/경제퀴즈", "/퀴즈목록", "/퀴즈등록", "/퀴즈삭제 과학 | 과학 문제" }))
                throw new Exception("퀘스트 조회·출석·퀴즈 별칭 수신 연결 오류");
            StaticVariable.AutoReboot = previousAuto;
            Console.WriteLine("PASS: periodic room recycle, draft deferral, retry, queue and receiver state preserved (no native input)");
            var dbType = typeof(Database);
            var refreshType = dbType.GetNestedType("ContentRefresh", BindingFlags.NonPublic);
            var snapshot = Activator.CreateInstance(refreshType, true);
            var flags = BindingFlags.Instance | BindingFlags.NonPublic;
            var incomingQuizzes = new List<Quiz> { new Quiz { Question = "새 문제", Category = "과학" }, new Quiz { Question = "경제 문제", Category = "경제" }, new Quiz { Question = "과학 문제", Category = " 과학 " } };
            refreshType.GetField("Commands", flags).SetValue(snapshot, new List<List<string>> { new List<string> { "/갱신시험", "새 응답" } });
            refreshType.GetField("Quizzes", flags).SetValue(snapshot, incomingQuizzes);
            refreshType.GetField("Topics", flags).SetValue(snapshot, new List<Topic> { new Topic { Title = "새 주제" } });
            var sourceType = typeof(System.Threading.Tasks.TaskCompletionSource<>).MakeGenericType(refreshType);
            var source = Activator.CreateInstance(sourceType);
            sourceType.GetMethod("SetResult").Invoke(source, new[] { snapshot });
            dbType.GetField("contentRefresh", flags).SetValue(Database.Instance, sourceType.GetProperty("Task").GetValue(source, null));
            var oldQuizzes = Database.Instance.CommonSenses;
            Database.Instance.CurrentAnswerIndex = 0;
            var apply = dbType.GetMethod("ApplyContentRefresh", flags);
            apply.Invoke(Database.Instance, null);
            if (!Database.Instance.Keywords.Contains("/갱신시험") || !object.ReferenceEquals(oldQuizzes, Database.Instance.CommonSenses))
                throw new Exception("진행 중 퀴즈 보호 또는 명령 갱신 오류");
            Database.Instance.CurrentAnswerIndex = -1;
            apply.Invoke(Database.Instance, null);
            if (!object.ReferenceEquals(incomingQuizzes, Database.Instance.CommonSenses)) throw new Exception("퀴즈 종료 후 갱신 누락");
            var failure = Activator.CreateInstance(sourceType);
            sourceType.GetMethod("SetException", new[] { typeof(Exception) }).Invoke(failure, new object[] { new Exception("시험 갱신 실패") });
            dbType.GetField("contentRefresh", flags).SetValue(Database.Instance, sourceType.GetProperty("Task").GetValue(failure, null));
            apply.Invoke(Database.Instance, null);
            if (!Database.Instance.ContentRefreshError.Contains("시험") || !object.ReferenceEquals(incomingQuizzes, Database.Instance.CommonSenses))
                throw new Exception("갱신 실패 시 기존 데이터 유지 오류");
            dbType.GetField("random", flags).SetValue(Database.Instance, System.Security.Cryptography.RandomNumberGenerator.Create());
            for (int i = 0; i < 200; i++)
            {
                Database.Instance.ResetCommonSense();
                string category = i % 2 == 0 ? "과학" : "경제";
                if (!Database.Instance.TrySetNextCommonSense(category) || Database.Instance.GetCurrentQuiz().Category.Trim() != category)
                    throw new Exception("분류별 출제 오류");
                var current = Database.Instance.GetCurrentQuiz();
                if (Database.Instance.TrySetNextCommonSense("경제") || !object.ReferenceEquals(current, Database.Instance.GetCurrentQuiz()))
                    throw new Exception("진행 중인 문제 덮어쓰기 오류");
            }
            Database.Instance.ResetCommonSense();
            if (Database.Instance.TrySetNextCommonSense("없는분류") || Database.Instance.GetCurrentQuiz() != null)
                throw new Exception("없는 분류에서 전체 문제 출제 오류");
            if (!Database.Instance.GetQuizCategoryHelp().Contains("/과학퀴즈 — 2문제")) throw new Exception("분류 목록 오류");
            if (!Database.Instance.TrySetNextCommonSense(null)) throw new Exception("전체 출제 오류");
            Database.Instance.ResetCommonSense();
            Quiz registered;
            string registrationError;
            if (!Quiz.TryParseRegistration("/퀴즈등록 과학 | 하 | 물의 화학식은? | H2O | 알파벳과 숫자 | 수소 두 개와 산소 한 개", out registered, out registrationError))
                throw new Exception("문제 등록 파싱 실패");
            if (!registered.ToRow().Select(Convert.ToString).SequenceEqual(new[] { "물의 화학식은?", "과학", "하", "H2O", "알파벳과 숫자", "수소 두 개와 산소 한 개" }))
                throw new Exception("시트 등록 열 순서 오류");
            foreach (string invalid in new[] { "/퀴즈등록", "/퀴즈등록 과학|하|문제|답|힌트", "/퀴즈등록 과학|하|문제||힌트|해설", "/퀴즈등록 과학|하|문제|답|힌트|해설|초과", "/퀴즈등록 과학|하|문제|답|힌트|해설\n다음줄" })
                if (Quiz.TryParseRegistration(invalid, out registered, out registrationError)) throw new Exception("잘못된 등록 입력 허용");
            if (Quiz.IsQuizCommand("/퀴즈등록") || !Quiz.IsQuizCommand("/과학퀴즈") || !Quiz.IsQuizCommand("/상식퀴즈"))
                throw new Exception("공개 퀴즈 호출과 등록 구분 실패");
            string deleteCategory, deleteQuestion;
            if (!Quiz.TryParseDeletion("/퀴즈삭제 과학 | 물의 화학식은?", out deleteCategory, out deleteQuestion) || deleteCategory != "과학" || deleteQuestion != "물의 화학식은?")
                throw new Exception("삭제 입력 파싱 오류");
            if (Quiz.TryParseDeletion("/퀴즈삭제 과학 |", out deleteCategory, out deleteQuestion) || Quiz.IsQuizCommand("/퀴즈삭제 과학 | 과학퀴즈"))
                throw new Exception("삭제 명령 분리 오류");
            var deletionRows = new List<List<string>> { new List<string> { "질문", "분류" }, new List<string> { "물의 화학식은?", "과학" }, new List<string> { "물의 화학식은?", "기타" } };
            if (Quiz.FindDeletionRow(deletionRows, "과학", "물의 화학식은?") != 1) throw new Exception("삭제 대상 행 오류");
            bool missingRejected = false, duplicateRejected = false;
            try { Quiz.FindDeletionRow(deletionRows, "과학", "물의"); } catch (InvalidOperationException) { missingRejected = true; }
            deletionRows.Add(new List<string> { "물의 화학식은?", "과학" });
            try { Quiz.FindDeletionRow(deletionRows, "과학", "물의 화학식은?"); } catch (InvalidOperationException) { duplicateRejected = true; }
            if (!missingRejected || !duplicateRejected) throw new Exception("부분 일치 또는 중복 삭제 허용 오류");
            Database.Instance.TrySetNextCommonSense("경제");
            var preservedQuiz = Database.Instance.GetCurrentQuiz();
            var deleteMethod = dbType.GetMethod("ApplyQuizDeletion", flags);
            if ((bool)deleteMethod.Invoke(Database.Instance, new object[] { "과학", "새 문제" }) || !object.ReferenceEquals(preservedQuiz, Database.Instance.GetCurrentQuiz()))
                throw new Exception("다른 문제 삭제 시 진행 중 문제 변경 오류");
            if (!(bool)deleteMethod.Invoke(Database.Instance, new object[] { "경제", "경제 문제" }) || Database.Instance.GetCurrentQuiz() != null)
                throw new Exception("진행 중인 삭제 문제 종료 오류");
            if (Database.Instance.TrySetNextCommonSense("경제")) throw new Exception("삭제된 문제 재출제 오류");
            Console.WriteLine("PASS: deletion routing, exact match, duplicate refusal, active quiz preservation/cancellation (no writes)");
            Console.WriteLine("PASS: registration command routing, validation and sheet column mapping (no writes)");
            Console.WriteLine("PASS: category commands, category-only selection, active quiz protection, unknown category and full pool");
            var executeRead = typeof(GoogleSheetHelper).GetMethod("ExecuteRead", BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(typeof(int));
            int attempts = 0;
            var waits = new List<int>();
            Func<int> temporaryTimeout = () => { attempts++; if (attempts < 3) throw new System.Threading.Tasks.TaskCanceledException(); return 42; };
            int recovered = (int)executeRead.Invoke(null, new object[] { temporaryTimeout, new Action<int>(waits.Add) });
            if (recovered != 42 || attempts != 3 || !waits.SequenceEqual(new[] { 500, 1500 })) throw new Exception("조회 재시도 복구 오류");
            attempts = 0;
            Func<int> persistentTimeout = () => { attempts++; throw new System.Threading.Tasks.TaskCanceledException(); };
            bool boundedFailure = false;
            try { executeRead.Invoke(null, new object[] { persistentTimeout, new Action<int>(ms => { }) }); }
            catch (TargetInvocationException error) { boundedFailure = error.InnerException is System.IO.IOException; }
            if (!boundedFailure || attempts != 3) throw new Exception("조회 시간 초과 횟수 제한 오류");
            attempts = 0;
            Func<int> invalidData = () => { attempts++; throw new FormatException("잘못된 데이터"); };
            try { executeRead.Invoke(null, new object[] { invalidData, new Action<int>(ms => { }) }); }
            catch (TargetInvocationException error) { if (!(error.InnerException is FormatException)) throw; }
            if (attempts != 1) throw new Exception("데이터 오류 재시도 금지 오류");
            var retainedQuizzes = Database.Instance.CommonSenses;
            var retainedCommands = Database.Instance.Keywords;
            var canceled = Activator.CreateInstance(sourceType);
            sourceType.GetMethod("SetCanceled", Type.EmptyTypes).Invoke(canceled, null);
            dbType.GetField("contentRefresh", flags).SetValue(Database.Instance, sourceType.GetProperty("Task").GetValue(canceled, null));
            apply.Invoke(Database.Instance, null);
            if (!Database.Instance.ContentRefreshError.Contains("시간 초과") || !object.ReferenceEquals(retainedQuizzes, Database.Instance.CommonSenses) || !object.ReferenceEquals(retainedCommands, Database.Instance.Keywords))
                throw new Exception("취소된 갱신의 상태 표시·캐시 유지 오류");
            var handled = Activator.CreateInstance(refreshType, true);
            refreshType.GetField("Error", flags).SetValue(handled, "콘텐츠 갱신 실패: 시험 시간 초과");
            var handledSource = Activator.CreateInstance(sourceType);
            sourceType.GetMethod("SetResult").Invoke(handledSource, new[] { handled });
            dbType.GetField("contentRefresh", flags).SetValue(Database.Instance, sourceType.GetProperty("Task").GetValue(handledSource, null));
            apply.Invoke(Database.Instance, null);
            if (!Database.Instance.ContentRefreshError.Contains("시험 시간 초과") || !object.ReferenceEquals(retainedQuizzes, Database.Instance.CommonSenses))
                throw new Exception("처리된 통신 실패가 기존 캐시를 덮어씀");
            Console.WriteLine("PASS: timeout recovery, bounded retries, no retry on invalid data, cancellation/error preserves cache");
            Console.WriteLine("PASS: background refresh apply, active quiz protection, failed refresh preserves cache");
            Console.WriteLine("PASS: Bot command queue, mention IDs, quiz, contribution, room isolation (no sends)");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static void Set(object target, string name, object value)
    { typeof(Bot).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(target, value); }
}
