using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace KakaotalkBot
{
    public partial class Bot
    {
        private ChatRosterSnapshot moderationRoster;
        private sealed class ModerationAccess
        {
            internal long At, ChatId, Session;
            internal bool BotCanManage;
            internal HashSet<long> Allowed, Protected;
            internal Dictionary<long, string> Names;
        }
        private volatile ModerationAccess moderationAccess;
        private readonly ConcurrentQueue<Action> moderationCompletions = new ConcurrentQueue<Action>();
        private void RefreshModerationAccess()
        {
            var db = Database.Instance; var roster = moderationRoster;
            long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            if (roster == null || SelectedRoom == null || roster.ChatId != SelectedRoom.ChatId || !roster.IsOpenGroup || db.Governance == null || !db.Governance.Enabled(roster.ChatId)) { moderationAccess = null; return; }
            var allowed = new HashSet<long>(); var protectedUsers = new HashSet<long> { OwnAuthorId };
            foreach (var member in roster.Members.Where(m => m.IsPresent))
            {
                ChatMemberState fresh; string error;
                if (!db.Operators.TryMember(roster.ChatId, member.UserId, now, out fresh, out error)) continue;
                bool judge = db.Governance.IsJudge(roster.ChatId, db.GovernanceUser(member.UserId), now);
                if (fresh.MemberType == 1 || fresh.MemberType == 4 || judge) allowed.Add(member.UserId);
                if (fresh.MemberType == 1 || fresh.MemberType == 4 || judge) protectedUsers.Add(member.UserId);
            }
            bool botCan = allowed.Contains(OwnAuthorId);
            if (BotIdentity.SuperUserId > 0) { allowed.Add(BotIdentity.SuperUserId); protectedUsers.Add(BotIdentity.SuperUserId); }
            moderationAccess = new ModerationAccess { At = roster.ObservedAt, ChatId = roster.ChatId, Session = Interlocked.Read(ref macroSession), BotCanManage = botCan,
                Allowed = allowed, Protected = protectedUsers, Names = roster.Members.Where(m => m.IsPresent).GroupBy(m => m.UserId).ToDictionary(g => g.Key, g => g.First().Nickname) };
        }
        private bool ModerationAllowed(long chat, long actor, long session, long target = 0, string nickname = null)
        {
            var access = moderationAccess; long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            return IsBotRunning && access != null && access.ChatId == chat && access.Session == session && session == Interlocked.Read(ref macroSession) && now >= access.At && now - access.At <= 30 && access.BotCanManage && access.Allowed.Contains(actor) &&
                (target == 0 || target != actor && !access.Protected.Contains(target) && access.Names.ContainsKey(target) && access.Names[target] == nickname && access.Names.Values.Count(n => n == nickname) == 1);
        }
        private bool TryModerationCommand(Command command)
        {
            string verb = OperatorCommandPolicy.Name(command.Keyword);
            if (verb != "/내보내기" && verb != "/내보내기해제" && verb != "/내보낸목록") return false;
            RefreshModerationAccess();
            long session = Interlocked.Read(ref macroSession), now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            if (command.ReceivedAt <= 0 || now < command.ReceivedAt || now - command.ReceivedAt > 15 || !ModerationAllowed(command.ChatId, command.AuthorId, session))
            { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "관리 권한 또는 최신 참여자 정보를 확인할 수 없습니다. 대법관·방장·부방장만 사용할 수 있습니다."); return true; }
            var db = Database.Instance; var store = db.Governance;
            string reason = "", nickname = null; long target = 0, kickedAt = 0;
            if (verb == "/내보내기")
            {
                if (!command.TryReadMentionText(verb, out target, out reason) || string.IsNullOrWhiteSpace(reason))
                { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "형식: /내보내기 @유저 사유"); return true; }
                nickname = db.ChatUserName(command.ChatId, target);
                if (!ModerationAllowed(command.ChatId, command.AuthorId, session, target, nickname))
                { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "대상이 보호 대상이거나, 현재 방에 없거나, 같은 닉네임이 여러 명입니다."); return true; }
            }
            else if (verb == "/내보내기해제")
            {
                string argument = command.Keyword.Substring(verb.Length).Trim(); int split = argument.IndexOf(' '), number;
                if (split < 1 || !int.TryParse(argument.Substring(0, split), out number) || string.IsNullOrWhiteSpace(argument.Substring(split + 1)))
                { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "먼저 /내보낸목록을 확인하고 /내보내기해제 번호 사유 형식으로 입력하세요."); return true; }
                var row = db.Governance.FindBan(command.ChatId, number);
                if (row == null) { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "목록 번호가 없습니다. /내보낸목록으로 다시 확인하세요."); return true; }
                target = db.GovernanceUserId(row.User); kickedAt = row.KickedAt; reason = argument.Substring(split + 1).Trim();
            }
            else if (command.Keyword.Trim() != verb)
            { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "형식: /내보낸목록"); return true; }
            if (reason.Length > 300 || reason.Any(char.IsControl)) { WindowsMacro.Instance.SendTextToChatroom(TargetWindow, "사유는 줄바꿈 없이 300자 이내로 입력하세요."); return true; }
            var room = SelectedRoom; string window = TargetWindow, key = command.ChatId + ":" + command.LogId;
            if (!db.Governance.QueueModeration(command.ChatId, key, verb, db.GovernanceUser(command.AuthorId), target > 0 ? db.GovernanceUser(target) : "", db.ForChat(reason), now)) return true;
            long selectedTarget = target, selectedAt = kickedAt;
            string selectedName = nickname;
            // 현재 명령 생성 범위에 한 단계로 등록되며, 실제 창 조작은 STA 소비자에서 실행됩니다.
            StaInputWorker.Instance.Invoke(() =>
            {
                string error = null; List<BannedMember> bans = null; ModerationMacro macro = null;
                try
                {
                    Func<bool> allowed = () => ModerationAllowed(command.ChatId, command.AuthorId, session, verb == "/내보내기" ? selectedTarget : 0, selectedName);
                    macro = new ModerationMacro(window, room.ProcessId, room.LinkId, allowed);
                    if (verb == "/내보내기") macro.Kick(selectedName);
                    else if (verb == "/내보내기해제") macro.Release(selectedTarget, selectedAt);
                    else bans = macro.ReadBans();
                }
                catch (Exception failure) { error = failure.GetBaseException().Message; }
                finally { if (macro != null) { try { macro.CloseOwnedDialogs(); } catch (Exception cleanup) { if (error == null) error = cleanup.GetBaseException().Message; } } }
                string result = error; var captured = bans;
                moderationCompletions.Enqueue(() =>
                {
                    if (captured != null && ReferenceEquals(db.Governance, store)) store.RecordBans(command.ChatId, captured.Select(b => new BanRecord { User = db.GovernanceUser(b.UserId), Name = db.ForChat(b.Name), KickedAt = b.KickedAt }));
                    store.FinishModeration(command.ChatId, key, db.ForChat(result));
                });
            });
            return true;
        }
    }
}
