using System;
using System.Collections.Generic;
using System.Linq;

namespace KakaotalkBot
{
    public partial class Bot
    {
        private static readonly HashSet<string> electionCommands = new HashSet<string> { "/선거", "/대법관", "/후보", "/선거결과", "/출마", "/출마취소", "/투표", "/투표취소", "/선거설정" };
        private const string DemocracyRoomName = "토론방 진실규명위원회";
        // 지정된 방에서는 활성화 전에도 설정 방법을 안내합니다.
        private string AppendGovernanceHelp(string text, long chatId)
        {
            var store = Database.Instance.Governance;
            bool enabled = store != null && store.Enabled(chatId);
            bool designated = SelectedRoom != null && SelectedRoom.ChatId == chatId &&
                SelectedRoom.IsOpenGroup && SelectedRoom.Name == DemocracyRoomName;
            if (!enabled && !designated) return text;
            return text + "\n\n[대법관 선거 · " + (enabled ? "사용 중" : "설정 필요") + "]" +
                "\n/선거 — 선거 상태·마감·다음 일정" +
                "\n/후보 — 후보 번호·공약 확인" +
                "\n/대법관 — 현재 대법관 목록" +
                "\n/선거결과 — 선거 결과 확인" +
                "\n/출마 공약 — 후보 등록·공약 수정 (300자 이내)" +
                "\n/출마취소 — 출마 철회" +
                "\n/투표 1 3 5 — 후보 번호를 공백으로 구분 (최대 10명)" +
                "\n/투표취소 — 내 투표 철회" +
                "\n다시 투표하면 이전 선택을 교체합니다. 투표 내용은 채팅에 공개됩니다." +
                "\n선거 시작 당시 참여자만 출마·투표할 수 있습니다." +
                "\n홀수달 1일 00시 시작·4일 00시 마감 (한국 시간), 임기 2개월" +
                "\n선거인 10% 이상 득표한 상위 최대 10명 선출, 결선은 남은 자리 수만큼 투표" +
                "\n\n[관리 · 대법관·방장·부방장]" +
                "\n/내보내기 @유저 사유 — 멘션한 참여자 내보내기" +
                "\n/내보낸목록 — 해제할 대상의 목록 번호 확인" +
                "\n/내보내기해제 번호 사유 — 해당 번호의 내보내기 해제" +
                "\n예: /내보내기해제 3 처리 완료" +
                "\n사유는 필수이며 줄바꿈 없이 300자 이내입니다." +
                "\n봇 계정에 방 관리 권한이 있어야 합니다." +
                "\n\n[선거 설정 · 방장]" +
                "\n/선거설정 켜기 — 자동 선거 활성화" +
                "\n/선거설정 끄기 — 선거 기능 비활성화";
        }
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
