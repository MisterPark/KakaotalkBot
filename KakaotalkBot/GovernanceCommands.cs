using System;
using System.Collections.Generic;
using System.Linq;

namespace KakaotalkBot
{
    public partial class Bot
    {
        private static readonly HashSet<string> electionCommands = new HashSet<string> { "/선거", "/대법관", "/후보", "/선거결과", "/출마", "/출마취소", "/투표", "/투표취소", "/선거설정" };
        private const string DemocracyRoomName = "토론방 진실규명위원회";
        private void ObserveGovernance(ChatRosterSnapshot roster)
        {
            var db = Database.Instance;
            if (db.Governance == null || OwnAuthorId <= 0 || !roster.IsOpenGroup) return;
            long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            // 경계 직전 명부를 먼저 사용한 뒤 새 관측으로 교체합니다.
            db.Governance.Pulse(roster.ChatId, now);
            var members = roster.Members.Where(m => m.IsPresent && m.UserId > 0 && m.UserId != OwnAuthorId).ToArray();
            moderationRoster = roster;
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
            store.Pulse(SelectedRoom.ChatId, DateTimeOffset.UtcNow.ToUnixTimeSeconds());
            string notice = store.PeekNotice(SelectedRoom.ChatId);
            if (notice == null) return;
            string destination = TargetWindow;
            if (StaInputWorker.Instance.TryBuildCommand(() => WindowsMacro.Instance.SendTextToChatroom(destination, notice)))
                store.TakeNotice(SelectedRoom.ChatId);
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
                else if (SelectedRoom.Name != DemocracyRoomName) result = "대법관 선거는 지정된 토론방에서만 설정할 수 있습니다.";
                else if (argument != "켜기" && argument != "끄기") result = "형식: /선거설정 켜기 또는 /선거설정 끄기";
                else if (OwnAuthorId <= 0) result = "봇 작성자 ID를 먼저 확인하세요.";
                else { db.Governance.Enable(command.ChatId, argument == "켜기", now); result = "선거 설정을 저장했습니다.\n" + db.Governance.Handle(command.ChatId, db.GovernanceUser(command.AuthorId), command.Nickname, "/선거", "", now); }
            }
            else result = db.Governance.Handle(command.ChatId, db.GovernanceUser(command.AuthorId), db.ForChat(command.Nickname), verb, db.ForChat(argument), now);
            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, result);
            return true;
        }
    }
}
