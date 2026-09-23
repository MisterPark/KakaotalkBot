using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KakaotalkBot
{
    public class Bot
    {
        public struct Chat
        {
            public long AuthorId;
            public string Nickname;
            public string Message;
        }
        public struct QuizAnswer
        {
            public long AuthorId;
            public string Nickname;
            public string Answer;
        }


        private string targetWindow = string.Empty;
        public string TargetWindow
        {
            get
            {
                return targetWindow; 
            }
            set 
            {
                targetWindow = value; 
                voiceRoomBot.TargetWindow = value;
            }
        }

        private bool isBotRunning = false;
        public bool IsBotRunning
        {
            get { return isBotRunning; }
            set
            {
                if (value) Start(); else Stop();
            }
        }

        private List<string> chatLog = new List<string>();
        private ChatDatabaseReceiver receiver;
        private DateTime nextOpenAttempt;
        private bool updating;
        private DateTime nextProcessingAttempt;
        private bool departureSavePending;
        public bool HasPendingUserSave { get { return departureSavePending || Database.Instance.HasPendingActivity; } }
        public string LastSendError { get; private set; }
        public string LastProcessingError { get; private set; }
        public long RequestedHistoryUserId { get; set; }
        internal Action<Command, long> OperatorHistoryRequested;
        internal long OperatorNotificationRoomId;
        internal string OperatorNotificationAccount;
        public ChatRoomInfo SelectedRoom { get; private set; }
        public long OwnAuthorId { get; set; }
        public bool HasReceiver { get { return receiver != null; } }
        public string ReceiveStatus { get { return receiver == null ? "채팅방을 선택해 주세요." :
            isBotRunning ? receiver.Status : (receiver.Completion.IsCompleted ? "수신 중지 · " + receiver.Status : "수신 종료 중…"); } }
        public bool IsReceiverStopping { get { return !isBotRunning && receiver != null && !receiver.Completion.IsCompleted; } }
        public System.Threading.Tasks.Task ReceiverCompletion { get { return receiver == null ? System.Threading.Tasks.Task.FromResult(0) : receiver.Completion; } }

        public void SelectRoom(ChatRoomInfo room)
        {
            if (isBotRunning || IsReceiverStopping) throw new InvalidOperationException("수신을 중지한 후 방을 변경해 주세요.");
            if (receiver != null) { receiver.Dispose(); receiver = null; }
            SelectedRoom = room;
            LastSendError = null;
            LastProcessingError = null;
            TargetWindow = room == null ? "" : room.Name;
            Reset();
        }

        private Queue<Command> commands = new Queue<Command>();
        private Queue<QuizAnswer> quizAnswers = new Queue<QuizAnswer>();

        private Random random;

        CustomTimer soliloquyTimer = new CustomTimer(300000);
        CustomTimer newsTimer = new CustomTimer(3600000);
        CustomTimer dbTimer = new CustomTimer(60000);

        private DateTime lastUpdate = DateTime.MinValue;

        private bool isCorrect = false;

        public Queue<Chat> DirectMessages = new Queue<Chat>();

        private enum PasswordStage { Verify, CheckOld, SetNew }
        private sealed class PasswordRequest
        {
            internal PasswordStage Stage;
            internal DateTime ExpiresAt;
        }
        private Dictionary<long, PasswordRequest> passwordRequests = new Dictionary<long, PasswordRequest>();
        private VoiceRoomBot voiceRoomBot;
        public VoiceRoomBot VoiceRoomBot { get { return voiceRoomBot; } }

        public Bot()
        {
            random = new Random(DateTime.Now.Millisecond);

            voiceRoomBot = new VoiceRoomBot();
        }

        public void Update()
        {
            if (IsBotRunning == false || updating) return;
            if (receiver != null && receiver.Completion.IsCompleted) { isBotRunning = false; return; }
            if (DateTime.UtcNow < nextProcessingAttempt) return;
            updating = true;
            try
            {
                UpdateInternally();
                if (IsBotRunning) voiceRoomBot.Update();
                nextProcessingAttempt = DateTime.MinValue;
                if (LastProcessingError != null && LastProcessingError.StartsWith("DB 수신 유지 · 처리 재시도 대기:", StringComparison.Ordinal)) LastProcessingError = null;
            }
            catch (Exception error) { DeferProcessing(error); }
            finally { updating = false; }
        }

        // 수신기는 계속 실행합니다. 저장 실패 등은 큐와 읽기 기준점을 초기화하지 않고 재시도합니다.
        internal void DeferProcessing(Exception error)
        {
            LastProcessingError = "DB 수신 유지 · 처리 재시도 대기: " + error.GetBaseException().Message;
            nextProcessingAttempt = DateTime.UtcNow.AddSeconds(5);
        }

        private void UpdateInternally()
        {
            Database.Instance.MaintainMonth();
            ProcessDatabaseMessages();
            ProcessCommand();
            if (!IsBotRunning) return;
            ProcessReset();

            if (soliloquyTimer.Check(Time.DeltaTime))
            {
                //ProcessComonBot();

                ProcessNextCommonSense();
            }

            if (newsTimer.Check(Time.DeltaTime))
            {
                ProcessNews();
            }

            if (dbTimer.Check(Time.DeltaTime))
            {
                ProcessUpdateDB();
            }

            ProcessQuiz();

        }

        public void Start()
        {
            if (isBotRunning) return;
            if (!Database.Instance.UserStorageReady) throw new InvalidOperationException(Database.Instance.UserStorageError ?? "사용자 DB를 먼저 정상적으로 불러와야 합니다.");
            Database.Instance.MaintainMonth();
            if (SelectedRoom == null) throw new InvalidOperationException("DB 목록에서 채팅방을 선택해 주세요.");
            if (string.IsNullOrWhiteSpace(TargetWindow)) throw new InvalidOperationException("답변을 보낼 채팅방 이름을 입력해 주세요.");
            if (IsReceiverStopping) throw new InvalidOperationException("이전 수신 작업을 종료하는 중입니다. 잠시 후 다시 시작해 주세요.");
            FlushPendingUserChanges(Database.Instance.UpdateUserTable);
            if (receiver != null) receiver.Dispose();
            Reset();
            LastSendError = null;
            receiver = new ChatDatabaseReceiver(SelectedRoom, OwnAuthorId);
            LastProcessingError = null;
            receiver.Start();
            isBotRunning = true;
        }

        public void Stop()
        {
            isBotRunning = false;
            if (receiver != null) receiver.Dispose();
            Reset();
            try { FlushPendingUserChanges(Database.Instance.UpdateUserTable); }
            catch (Exception error) { LastProcessingError = "사용자 활동·이력 저장 실패: " + error.GetBaseException().Message + " · 다시 시작하면 저장을 재시도합니다."; }
        }
        public void Reset()
        {
            nextProcessingAttempt = DateTime.MinValue;
            Database.Instance.Operators.ResetLive();
            chatLog = new List<string>();

            commands.Clear();
            quizAnswers.Clear();
            isCorrect = false;
            passwordRequests.Clear();
            soliloquyTimer = new CustomTimer(300000);
            newsTimer = new CustomTimer(3600000);
            dbTimer = new CustomTimer(60000);
        }

        private void ProcessDatabaseMessages()
        {
            if (receiver == null) return;
            ChatMessage chat;
            bool changed = false;
            for (int i = 0; i < 128 && commands.Count < 256 && receiver.TryRead(out chat); i++)
            {
                changed |= HandleIncomingMessage(chat);
            }
            // 입퇴장 응답을 보내기 전에 같은 수신 묶음의 횟수를 저장한다. 60초 주기를 기다리지 않는다.
            if (departureSavePending) FlushPendingUserChanges(Database.Instance.UpdateUserTable);
            if (changed && Form1.Instance != null) Form1.Instance.UpdateChatLog(string.Join("\r\n", chatLog));
        }

        private void FlushPendingUserChanges(Action save)
        {
            if (!HasPendingUserSave) return;
            try { save(); }
            catch { ChatEventDiagnostics.Write("save-failed"); throw; }
            // 저장이 실패하면 대기 표시를 유지해 중지·재시작 시 재시도한다.
            departureSavePending = false;
            Database.Instance.ConfirmActivitySaved();
            ChatEventDiagnostics.Write("saved");
        }

        private bool HandleIncomingMessage(ChatMessage chat)
        {
            if (chat.RoleChanged)
            {
                if (SelectedRoom != null && chat.ChatId == SelectedRoom.ChatId)
                    Database.Instance.Operators.Invalidate(chat.ChatId, chat.AuthorId, chat.LogId, chat.ExpectedMemberType);
                return false;
            }
            if (chat.Roster != null)
            {
                if (SelectedRoom != null && chat.ChatId == SelectedRoom.ChatId) Database.Instance.ObserveRoster(chat.Roster);
                return false;
            }
            if (chat.IsDirect)
            {
                // 암호 메시지는 화면이나 채팅 로그에 남기지 않습니다.
                HandleDirectMessage(new Chat { AuthorId = chat.AuthorId, Nickname = chat.Nickname, Message = chat.Message });
                return false;
            }
            if (chat.ChatId == OperatorNotificationRoomId && SelectedRoom != null && SelectedRoom.AccountPath == OperatorNotificationAccount) return false;
            if (SelectedRoom == null || chat.ChatId != SelectedRoom.ChatId || chat.AuthorId <= 0) return false;
            if (chat.RoomEvent == ChatRoomEvent.Join)
                Database.Instance.Operators.Invalidate(chat.ChatId, chat.AuthorId, chat.LogId, -4);
            if (chat.RoomEvent == ChatRoomEvent.Leave || chat.RoomEvent == ChatRoomEvent.Kick)
            {
                Database.Instance.Operators.Invalidate(chat.ChatId, chat.AuthorId, chat.LogId, -3);
                if (!Database.Instance.RecordDeparture(chat.AuthorId, chat.Nickname, chat.ChatId, chat.LogId, chat.RoomEvent == ChatRoomEvent.Kick, chat.SendAt)) return false;
                departureSavePending = true;
                ChatEventDiagnostics.Write("counted", chat);
            }
            var incomingUser = Database.Instance.AddUser(chat.AuthorId, chat.Nickname);
            if (Database.Instance.ObserveMessage(chat)) departureSavePending = true;
            // 봇 계정도 등록·방별 상태를 저장하되 자기 출력은 명령과 퀴즈에 넣지 않습니다.
            if (chat.IsOwn) return false;
            if (string.IsNullOrWhiteSpace(chat.Nickname)) chat.Nickname = Database.Instance.ChatUserName(chat.ChatId, chat.AuthorId);
            string message = (chat.Message ?? "").Trim();
            chatLog.Add("[" + chat.Nickname + "] " + message);
            if (chatLog.Count > 300) chatLog.RemoveAt(0);
            if (chat.EventCommand != null)
                ProcessKeyword(chat.Nickname, chat.EventCommand, chat.AuthorId, chat.LogId, chat.ChatId);
            else
            {
                ProcessQuizAnswer(chat.AuthorId, chat.Nickname, message);
                if (OperatorCommandPolicy.RequiresOperator(message) || OperatorCommandPolicy.Name(message) == "/통계" ||
                    OperatorCommandPolicy.Name(message) == "/월간랭킹" || OperatorCommandPolicy.Name(message) == "/채팅랭킹" ||
                    OperatorCommandPolicy.Name(message) == "/레벨랭킹" || OperatorCommandPolicy.Name(message) == "/랭킹" || Database.Instance.Keywords.Any(k => message.StartsWith(k)))
                    ProcessKeyword(chat.Nickname, message, chat.AuthorId, chat.LogId, chat.ChatId, chat.Mentions);
            }
            return true;
        }
        private void ProcessKeyword(string nickname, string message, long authorId = 0, long logId = 0, long chatId = 0,
            IReadOnlyList<ChatMention> mentions = null)
        {
            //SendTextToChatroom(TargetWindow, message);

            Command command = new Command();
            command.Nickname = nickname;
            command.Keyword = message;
            command.AuthorId = authorId;
            command.LogId = logId;
            command.ChatId = chatId;
            command.ReceivedAt = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            command.Mentions = mentions;
            commands.Enqueue(command);
        }

        private void ProcessQuizAnswer(long authorId, string nickname, string answer)
        {
            QuizAnswer quizAnswer = new QuizAnswer();
            quizAnswer.AuthorId = authorId;
            quizAnswer.Nickname = nickname;
            quizAnswer.Answer = answer;
            quizAnswers.Enqueue(quizAnswer);

            if(Database.Instance.FindUser(authorId,out User user))
            {
                user.Contribution += 1;
            }
        }

        private void ProcessCommand()
        {
            if (commands.Count == 0) return;


            // 전송 대상 창이 준비되지 않았으면 명령을 큐에 남겨 둡니다.
            if (!WindowsMacro.Instance.IsChatRoomOpen(TargetWindow))
            {
                if (DateTime.UtcNow >= nextOpenAttempt)
                {
                    nextOpenAttempt = DateTime.UtcNow.AddSeconds(3);
                    WindowsMacro.Instance.OpenChatRoom(TargetWindow);
                }
                return;
            }
            Command command = commands.Dequeue();
            if (command.AuthorId <= 0 || SelectedRoom == null || command.ChatId != SelectedRoom.ChatId) return;
            if (OperatorCommandPolicy.RequiresOperator(command.Keyword))
            {
                string permissionError;
                long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                if (!SelectedRoom.IsOpenGroup)
                    permissionError = "운영진 명령은 오픈채팅 그룹방에서만 사용할 수 있습니다.";
                else if (command.ReceivedAt <= 0 || now < command.ReceivedAt || now - command.ReceivedAt > 15)
                    permissionError = "명령 처리 대기 시간이 길어 취소했습니다. 다시 입력해 주세요.";
                else if (Database.Instance.Operators.Check(command.ChatId, command.AuthorId, now, out permissionError))
                    permissionError = null;
                if (permissionError != null)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, permissionError);
                    return;
                }
            }
            //string answer = GetAnswer(command.Keyword);
            string operation = OperatorCommandPolicy.Name(command.Keyword);
            if (operation == "/이력" || operation == "/메모")
            {
                long target; string body;
                if (!command.TryReadMentionText(operation, out target, out body) || (operation == "/이력" && body.Length != 0) || (operation == "/메모" && body.Length == 0))
                { LastProcessingError = "형식: /이력 @유저 또는 /메모 @유저 내용 (실제 멘션 필요)"; return; }
                if (operation == "/메모") Database.Instance.AddOperatorMemo(command.ChatId, command.AuthorId, target, body, command.LogId);
                RequestedHistoryUserId = target;
                if (operation == "/이력" && OperatorHistoryRequested != null) OperatorHistoryRequested(command, target);
                return;
            }
            if (operation == "/레벨랭킹")
            {
                if (command.Keyword.Trim() != "/레벨랭킹")
                { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "형식: /레벨랭킹 (현재 누적 경험치 기준)"); return; }
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, Database.Instance.LevelRanking(command.ChatId));
                return;
            }
            if (operation == "/통계" || operation == "/월간랭킹" || operation == "/채팅랭킹" || operation == "/랭킹")
            {
                string month = command.Keyword.Substring(operation.Length).Trim();
                if (month.Length == 0) month = OperationsStore.Month(DateTimeOffset.UtcNow.ToUnixTimeSeconds());
                DateTime parsed;
                if (!DateTime.TryParseExact(month, "yyyy-MM", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out parsed))
                { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "형식: " + operation + " [YYYY-MM]"); return; }
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, operation == "/통계" ? Database.Instance.RoomStatistics(command.ChatId, month) : Database.Instance.Operations.Ranking(command.ChatId, month, operation == "/랭킹"));
                return;
            }

            //if (string.IsNullOrEmpty(answer) == false)
            //{
            //    SendTextToChatroom(TargetWindow, $"{answer}");
            //}

            if (command.Keyword == "/?" || command.Keyword == "/명령어" || command.Keyword == "/훈장" || command.Keyword == "/공지사항" || command.Keyword == "/패치노트" || command.Keyword == "/서브방")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
                }
            }
            else if (command.Keyword == "/입장")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SmartString.CurrentNickname = command.Nickname;
                    string parsedAnswer = SmartString.Parse(answer);
                    try
                    {
                        if (SelectedRoom == null || command.ChatId != SelectedRoom.ChatId || command.AuthorId <= 0)
                            throw new InvalidOperationException("입장 명령의 채팅방 또는 작성자 ID가 올바르지 않습니다.");
                        WindowsMacro.Instance.SendMentionToChatroom(TargetWindow, command.AuthorId, command.Nickname, parsedAnswer, SelectedRoom.ProcessId, Database.Instance.MentionFallbackNames(command.ChatId, command.AuthorId));
                    }
                    catch (Exception error)
                    {
                        // 실패한 멘션에 일반 본문을 덧붙이거나 같은 명령을 재전송하지 않는다.
                        LastSendError = "DB 수신 유지 · 해당 멘션 전송 실패: " + error.GetBaseException().Message;
                    }
                }
            }
            else if (command.Keyword == "/퇴장")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SmartString.CurrentNickname = command.Nickname;
                    string parsedAnswer = SmartString.Parse(answer);
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, parsedAnswer);
                }
            }
            else if (command.Keyword == "/출첵")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    if (Database.Instance.CheckAttendance(command.AuthorId, command.Nickname))
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"이미 출석한 유저입니다.");
                    }
                    else
                    {
                        if (Database.Instance.FindUser(command.AuthorId, out User user))
                        {
                            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{command.Nickname}]님이 {answer}\n+10포인트\n(현재 포인트: {user.Point})");
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{command.Nickname}]님이 {answer}\n+10포인트");
                        }


                    }
                }
            }
            else if (command.Keyword.StartsWith("/조회"))
            {
                long targetId = command.AuthorId;
                string targetNickname = command.Nickname;
                int ignored;
                if (command.Keyword != "/조회" && !command.TryReadTarget("/조회", false, out targetId, out ignored, out targetNickname))
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "형식: /조회 또는 /조회 @사용자\n다른 사용자는 카카오톡 멘션 후보에서 직접 선택해 주세요.");
                    return;
                }
                var user = Database.Instance.GetOrAddUser(targetId, targetNickname);
                int total = Database.Instance.GetTotalContribution();
                float contribution = total == 0 ? 0 : user.Contribution * 100f / total;
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"=====[유저조회]=====\n닉네임: {Database.Instance.ChatUserName(command.ChatId, targetId, targetNickname)}\n레벨: {user.Level}\n경험치: {user.Experience} (다음 레벨 누적 {user.NextLevelExperience})\n{Database.Instance.PersonalRankings(command.ChatId, targetId, OperationsStore.Month(DateTimeOffset.UtcNow.ToUnixTimeSeconds()))}\n포인트: {user.Point}\n인기도: {user.Popularity}\n채팅 기여도: {contribution:F2}%\n퇴장 횟수: {user.LeaveCount}\n강퇴 횟수: {user.KickCount}\n{Database.Instance.DescribeActivity(command.ChatId, targetId)}\n=================");
            }
            else if (command.Keyword.StartsWith("/랭킹"))
            {
                if (command.Keyword == "/랭킹")
                {
                    string answer = Database.Instance.GetAnswer(command.Keyword);

                    if (string.IsNullOrEmpty(answer)) return;

                    StringBuilder sb = new StringBuilder();
                    int beforeRank = 1;
                    int beforePop = 0;
                    List<User> rank = Database.Instance.GetPopularityRank();

                    sb.AppendLine(answer);
                    for (int i = 0; i < Math.Min(20, rank.Count); i++)
                    {
                        int currentPop = rank[i].Popularity;
                        if (currentPop != beforePop)
                        {
                            beforeRank = i + 1;
                        }

                        string emoji = string.Empty;
                        if (beforeRank == 1)
                        {
                            emoji = "🥇";
                        }
                        else if (beforeRank == 2)
                        {
                            emoji = "🥈";
                        }
                        else if (beforeRank == 3)
                        {
                            emoji = "🥉";
                        }

                        sb.AppendLine($"{emoji}{beforeRank}위 {Database.Instance.ChatUserName(SelectedRoom == null ? 0 : SelectedRoom.ChatId, rank[i].UserId)} {rank[i].Popularity}");

                        beforePop = currentPop;
                    }

                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/토론"))
            {
                if (command.Keyword == "/토론")
                {
                    string answer = Database.Instance.GetAnswer(command.Keyword);

                    if (string.IsNullOrEmpty(answer)) return;

                    StringBuilder sb = new StringBuilder();

                    List<Topic> topics = Database.Instance.GetOrderedTopics();
                    string currentCategory = string.Empty;
                    int count = 0;

                    sb.AppendLine(answer);
                    for (int i = 0; i < topics.Count; i++)
                    {
                        if (currentCategory != topics[i].Category)
                        {
                            sb.AppendLine();
                            sb.AppendLine($"{topics[i].Category}");
                            currentCategory = topics[i].Category;
                            count = 0;
                        }
                        sb.AppendLine($"{count + 1}. {topics[i].Title}({topics[i].CreatedAt})");
                        sb.AppendLine($"(토론왕👑: {topics[i].Winner})");
                        count++;
                    }

                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/좋아") || command.Keyword.StartsWith("/싫어"))
            {
                bool increase = command.Keyword.StartsWith("/좋아");
                string verb = increase ? "/좋아" : "/싫어";
                long targetId;
                int amount;
                string targetNickname;
                if (!command.TryReadTarget(verb, true, out targetId, out amount, out targetNickname))
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "형식: " + verb + " @사용자 [양의 숫자]\n카카오톡 멘션 후보에서 한 명을 선택해 주세요. 숫자를 생략하면 10포인트입니다.");
                    return;
                }
                var author = Database.Instance.AddUser(command.AuthorId, command.Nickname);
                var target = Database.Instance.GetOrAddUser(targetId, targetNickname);
                string error;
                if (!Database.Instance.ChangePopularity(command.AuthorId, targetId, amount, increase, out error, command.ChatId, command.LogId))
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, error);
                    return;
                }
                string action = increase ? "👍좋아요" : "👎싫어요";
                string direction = increase ? "상승" : "하락";
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{Database.Instance.ChatUserName(command.ChatId, command.AuthorId, command.Nickname)}]님이 [{Database.Instance.ChatUserName(command.ChatId, targetId, targetNickname)}]님에게 {action}.\n인기도 {amount}점 {direction}\n현재 인기도: {target.Popularity}");
            }
            else if (command.Keyword.StartsWith("/정치뉴스"))
            {
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{News.PoliticsTop6}");
            }
            else if (command.Keyword == "/상식퀴즈")
            {
                ProcessCommonSense();
            }
            else if (command.Keyword == "/암호검증")
            {
                receiver.WatchDirect(command.AuthorId);
                ProcessVerifyPassword(command.AuthorId);
            }
            else if (command.Keyword == "/암호변경")
            {
                receiver.WatchDirect(command.AuthorId);
                ProcessChangePassword(command.AuthorId);
            }
            else if (command.Keyword == "/채팅랭킹")
            {
                ProcessContribution();
            }
            else
            {
                string[] answers = Database.Instance.GetAnswers(command.Keyword);
                if (answers == null || answers.Length == 0)
                {
                    return;
                }

                int rand = random.Next(0, answers.Length);
                string answer = answers[rand];

                if (string.IsNullOrEmpty(answer) == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
                }
            }

        }

        private void ProcessCommonSense()
        {
            string answer = Database.Instance.GetCommonSenseText();
            if (string.IsNullOrEmpty(answer))
            {
                //double left = soliloquyTimer.TimeLeft * 0.001;

                //WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"다음 퀴즈를 준비하고 있습니다.\n남은 시간: {left}초");
                Database.Instance.SetNextCommonSense();
                answer = Database.Instance.GetCommonSenseText();
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
                soliloquyTimer.Reset();
            }
            else
            {
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
            }
        }

        private void ProcessReset()
        {
            if (lastUpdate.Day != DateTime.Now.Day)
            {
                lastUpdate = DateTime.Now;
                Database.Instance.ResetAttendance();
            }
        }

        private void ProcessNextCommonSense()
        {
            Quiz quiz = Database.Instance.GetCurrentQuiz();
            if (quiz != null)
            {
                if (isCorrect == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정답자가 없습니다.\n정답: {quiz.Answer}\n해설: {quiz.Explanation}");
                }
            }

            Database.Instance.ResetCommonSense();
            //Database.Instance.SetNextCommonSense();
            //ProcessCommonSense();
        }

        private void ProcessNews()
        {
            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{News.PoliticsTop6}");
        }

        private void ProcessUpdateDB()
        {
            Database.Instance.UpdateCommands();
            Database.Instance.UpdateUserTable();
            Database.Instance.UpdateCommonSenses();
            Database.Instance.UpdateTopic();
            News.Update();
        }

        private void ProcessQuiz()
        {
            Quiz quiz = Database.Instance.GetCurrentQuiz();
            if (quiz == null)
            {
                return;
            }

            while (quizAnswers.Count != 0)
            {
                QuizAnswer quizAnswer = quizAnswers.Dequeue();

                string[] answers = quiz.Answer.Split(new string[] {", " }, StringSplitOptions.RemoveEmptyEntries);

                string[] userAnswers = quizAnswer.Answer.Split(new string[] { ", "}, StringSplitOptions.RemoveEmptyEntries);

                

                if (SmartString.SameElements(quiz.Answer, quizAnswer.Answer))
                {
                    if (Database.Instance.FindUser(quizAnswer.AuthorId, out User a))
                    {
                        int point = 0;
                        if (quiz.Difficulty == "최상")
                        {
                            point = 5;
                        }
                        else if (quiz.Difficulty == "상")
                        {
                            point = 4;
                        }
                        else if (quiz.Difficulty == "중")
                        {
                            point = 3;
                        }
                        else if (quiz.Difficulty == "하")
                        {
                            point = 2;
                        }
                        else
                        {
                            point = 1;
                        }

                        a.Point += point;
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"💡정답자: {quizAnswer.Nickname}\n💬정답: {quiz.Answer}\n📜해설: {quiz.Explanation}\n+{point} 포인트 득점!!👍\n 현재 포인트: {a.Point}");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"💡정답자: {quizAnswer.Nickname}\n💬정답: {quiz.Answer}\n📜해설: {quiz.Explanation}");
                    }

                    Database.Instance.CurrentAnswerIndex = -1;
                    break;
                }
            }

            quizAnswers.Clear();

        }

        private void ProcessVerifyPassword(long authorId)
        {
            passwordRequests[authorId] = new PasswordRequest { Stage = PasswordStage.Verify, ExpiresAt = DateTime.UtcNow.AddMinutes(5) };
            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "[암호검증 방법]\n5분 안에 봇에게 1대1 메시지로 암호를 보내세요. 요청한 사용자 ID와 일치하는 대화만 처리합니다.");
        }

        private void ProcessChangePassword(long authorId)
        {
            passwordRequests[authorId] = new PasswordRequest { Stage = PasswordStage.CheckOld, ExpiresAt = DateTime.UtcNow.AddMinutes(5) };
            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "[암호변경]\n5분 안에 봇에게 1대1 메시지로 현재 암호를 보내세요. 요청한 사용자 ID와 일치하는 대화만 처리합니다.");
        }

        private void HandleDirectMessage(Chat chat)
        {
            string response = ProcessDirectMessage(chat);
            if (response != null) WindowsMacro.Instance.SendTextToChatroom(TargetWindow, response);
        }

        // 인증 대기 상태와 사용자 조회는 모두 DB가 전달한 실제 작성자 ID로만 연결한다.
        private string ProcessDirectMessage(Chat chat)
        {
            PasswordRequest request;
            if (chat.AuthorId <= 0 || !passwordRequests.TryGetValue(chat.AuthorId, out request)) return null;
            if (DateTime.UtcNow >= request.ExpiresAt) { passwordRequests.Remove(chat.AuthorId); return null; }
            var user = Database.Instance.GetOrAddUser(chat.AuthorId, chat.Nickname);
            if (request.Stage == PasswordStage.SetNew)
            {
                passwordRequests.Remove(chat.AuthorId);
                if (string.IsNullOrEmpty(chat.Message)) return "빈 암호로 변경할 수 없습니다.";
                user.Password = chat.Message;
                return "[" + (string.IsNullOrWhiteSpace(chat.Nickname) ? "이름 미확인 사용자" : chat.Nickname) + "]\n암호가 변경되었습니다.";
            }
            if (user.Password != chat.Message)
            {
                passwordRequests.Remove(chat.AuthorId);
                return "[" + (string.IsNullOrWhiteSpace(chat.Nickname) ? "이름 미확인 사용자" : chat.Nickname) + "]\n암호가 일치하지 않습니다.";
            }
            if (request.Stage == PasswordStage.Verify)
            {
                passwordRequests.Remove(chat.AuthorId);
                return "[" + (string.IsNullOrWhiteSpace(chat.Nickname) ? "이름 미확인 사용자" : chat.Nickname) + "]\n암호가 일치합니다.";
            }
            request.Stage = PasswordStage.SetNew;
            return "[" + (string.IsNullOrWhiteSpace(chat.Nickname) ? "이름 미확인 사용자" : chat.Nickname) + "]\n변경할 암호를 1대1 메시지로 보내주세요.";
        }
        private void ProcessContribution()
        {
            StringBuilder sb = new StringBuilder();
            int beforeRank = 1;
            int beforePop = 0;
            List<User> rank = Database.Instance.GetContributionRank();
            int totalContribution = Database.Instance.GetTotalContribution();
            float totalContribution2 = totalContribution == 0 ? 1 : totalContribution;

            sb.AppendLine("채팅 기여도 랭킹");
            for (int i = 0; i < Math.Min(20, rank.Count); i++)
            {
                float userContribution = rank[i].Contribution / totalContribution2 * 100;
                int currentPop = rank[i].Contribution;
                if (currentPop != beforePop)
                {
                    beforeRank = i + 1;
                }

                string emoji = string.Empty;
                if (beforeRank == 1)
                {
                    emoji = "🥇";
                }
                else if (beforeRank == 2)
                {
                    emoji = "🥈";
                }
                else if (beforeRank == 3)
                {
                    emoji = "🥉";
                }

                sb.AppendLine($"{emoji}{beforeRank}위 | {Database.Instance.ChatUserName(SelectedRoom == null ? 0 : SelectedRoom.ChatId, rank[i].UserId)} | {userContribution:F2}% | {currentPop}");

                beforePop = currentPop;
            }

            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, sb.ToString());
        }
    }
}
