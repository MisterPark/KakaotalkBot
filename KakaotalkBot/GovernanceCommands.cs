using System;
using System.Collections.Generic;
using System.Linq;

namespace KakaotalkBot
{
    public partial class Bot
    {
        private static readonly HashSet<string> electionCommands = new HashSet<string> { "/선거", "/대법관", "/후보", "/선거결과", "/선거기록", "/출마", "/출마취소", "/투표", "/투표취소", "/선거설정", "/선거시작", "/보궐선거" };
        internal static bool IsGovernanceCommand(string message)
        {
            string name = OperatorCommandPolicy.Name(message);
            return electionCommands.Contains(name) || IsModerationCommand(name);
        }
        private static bool IsModerationCommand(string name)
        {
            return name == "/내보내기" || name == "/내보내기해제" || name == "/내보낸목록";
        }
        private long CurrentOwnAuthorId
        {
            get
            {
                if (OwnAuthorId > 0) return OwnAuthorId;
                var roster = moderationRoster;
                long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                if (SelectedRoom == null || roster == null || roster.ChatId != SelectedRoom.ChatId || roster.LinkId != SelectedRoom.LinkId ||
                    now < roster.ObservedAt || now - roster.ObservedAt > 15) return 0;
                return roster.Members.Any(m => m.IsPresent && m.UserId == roster.OwnAuthorId) ? roster.OwnAuthorId : 0;
            }
        }
        private void ObserveGovernance(ChatRosterSnapshot roster)
        {
            var db = Database.Instance;
            // 방 관리 명부는 선거와 봇 작성자 ID의 수동 설정 여부에 관계없이 갱신합니다.
            moderationRoster = roster;
            long ownId = CurrentOwnAuthorId;
            if (db.Governance == null || ownId <= 0 || !roster.IsOpenGroup) return;
            long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            // 경계 직전 명부를 먼저 사용한 뒤 새 관측으로 교체합니다.
            db.Governance.Pulse(roster.ChatId, now);
            var members = roster.Members.Where(m => m.IsPresent && m.UserId > 0 && m.UserId != ownId).ToArray();
            db.Governance.Observe(roster.ChatId, new GovernanceRoster { At = roster.ObservedAt,
                Present = new HashSet<string>(members.Select(m => db.GovernanceUser(m.UserId))),
                Names = members.GroupBy(m => db.GovernanceUser(m.UserId)).ToDictionary(g => g.Key, g => g.First().Nickname) });
        }
        private void TickGovernance()
        {
            Action completion;
            while (moderationCompletions.TryDequeue(out completion)) completion();
            RefreshModerationAccess();
            var store = Database.Instance.Governance;
            if (store == null || SelectedRoom == null) return;
            long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            store.Pulse(SelectedRoom.ChatId, now);
            string notice = store.PeekNotice(SelectedRoom.ChatId);
            if (notice == null) return;
            string destination = TargetWindow;
            var room = SelectedRoom;
            long session = System.Threading.Interlocked.Read(ref macroSession);
            Func<bool> current = () => IsBotRunning && session == System.Threading.Interlocked.Read(ref macroSession) && ReferenceEquals(room, SelectedRoom) && ReferenceEquals(store, Database.Instance.Governance);
            // 창이 닫혔으면 알림을 보존하고 먼저 재열기를 요청합니다.
            if (!WindowsMacro.Instance.IsChatRoomOpen(destination))
            {
                if (DateTime.UtcNow >= nextOpenAttempt)
                {
                    nextOpenAttempt = DateTime.UtcNow.Add(RoomOpenRetryInterval);
                    StaInputWorker.Instance.TryBuildCommand(() => WindowsMacro.Instance.OpenChatRoom(destination, room.ProcessId), current);
                }
                return;
            }
            if (StaInputWorker.Instance.TryBuildCommand(() => WindowsMacro.Instance.SendTextToChatroom(destination, notice), () =>
                current() && (!GovernanceStore.IsVotingReminder(notice) || DateTimeOffset.UtcNow.ToUnixTimeSeconds() < now + 3600) &&
                store.CanSendNotice(room.ChatId, notice, DateTimeOffset.UtcNow.ToUnixTimeSeconds())))
                store.TakeNotice(room.ChatId);
        }
        private bool TryGovernanceCommand(Command command)
        {
            string verb = OperatorCommandPolicy.Name(command.Keyword);
            if (TryModerationCommand(command)) return true;
            if (!electionCommands.Contains(verb)) return false;
            var db = Database.Instance;
            string result;
            long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            string argument = command.Keyword.Substring(verb.Length).Trim();
            if (db.Governance == null || SelectedRoom == null || !SelectedRoom.IsOpenGroup) result = "선거 DB 또는 오픈채팅방 정보를 확인하세요.";
            else if (verb == "/선거설정")
            {
                string error;
                if (!db.Operators.CheckQuizIssuer(command.ChatId, command.AuthorId, BotIdentity.SuperUserId, now, command.ReceivedAt, out error)) result = error;
                else if (argument != "켜기" && argument != "끄기") result = "형식: /선거설정 켜기 또는 /선거설정 끄기";
                else if (argument == "켜기" && CurrentOwnAuthorId <= 0) result = "봇 계정 ID를 자동 확인 중입니다. 잠시 후 다시 시도하거나 설정의 ‘본인 작성자 ID’를 확인하세요.";
                else { db.Governance.Enable(command.ChatId, argument == "켜기", now); result = "선거 설정을 저장했습니다.\n" + db.Governance.Handle(command.ChatId, db.GovernanceUser(command.AuthorId), command.Nickname, "/선거", "", now); }
            }
            else if (verb == "/선거시작" || verb == "/보궐선거")
            {
                string error;
                if (!db.Operators.CheckQuizIssuer(command.ChatId, command.AuthorId, BotIdentity.SuperUserId, now, command.ReceivedAt, out error)) result = error;
                else if (argument.Length != 0) result = "형식: " + verb;
                else if (CurrentOwnAuthorId <= 0) result = "봇 계정 ID를 자동 확인 중입니다. 잠시 후 다시 시도하세요.";
                else result = db.Governance.OpenElection(command.ChatId, verb == "/보궐선거", now);
            }
            else result = db.Governance.Handle(command.ChatId, db.GovernanceUser(command.AuthorId), db.ForChat(command.Nickname), verb, db.ForChat(argument), now);
            foreach (string page in GovernanceStore.Messages(result)) WindowsMacro.Instance.SendTextToChatroom(TargetWindow, page);
            return true;
        }
    }
}
