using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace KakaotalkBot
{
    internal sealed class PersonalReplyService
    {
        private readonly ChatDatabaseReceiver receiver;
        private readonly PersonalReplyAi ai;
        internal readonly PersonalReplyCoordinator Replies;
        private Task completion;
        private readonly CancellationTokenSource catalogCancellation = new CancellationTokenSource();
        private readonly Task<System.Collections.Generic.List<PersonalEmoticon>> catalog;
        internal readonly bool Automatic;
        internal bool IsStopping { get { return completion != null; } }
        internal Task Completion { get { return completion ?? Task.FromResult(0); } }
        internal string Status { get { return receiver.Status + (catalog.IsFaulted || catalog.IsCanceled ? " · 이모티콘 목록 미확인(텍스트·공감 사용)" : catalog.Status == TaskStatus.RanToCompletion ? " · 이모티콘 " + catalog.Result.Count + "개" : " · 이모티콘 목록 읽는 중"); } }
        internal PersonalReplyService(ChatRoomInfo[] rooms, PersonalReplySettings settings, bool automatic)
        {
            Automatic = automatic;
            receiver = new ChatDatabaseReceiver(rooms, settings.ImportHistory);
            ai = new PersonalReplyAi(settings.GetApiKey(), settings.Model);
            catalog = Task.Run(() => PersonalReplyActions.Load(rooms[0], catalogCancellation.Token), catalogCancellation.Token);
            Replies = new PersonalReplyCoordinator(rooms, settings, ai, automatic, Send, emoticons: () =>
                catalog.Status == TaskStatus.RanToCompletion ? (System.Collections.Generic.IReadOnlyList<PersonalEmoticon>)catalog.Result : new PersonalEmoticon[0]);
            receiver.Start();
        }
        internal void Tick()
        {
            if (IsStopping) return;
            ChatMessage message;
            // 전송 처리 전에 수신을 읽어, AI 작성 중 도착한 본인 메시지를 우선합니다.
            int count = 0;
            while (count++ < 4096 && receiver.TryRead(out message)) Replies.Receive(message);
            if (receiver.Completion.IsCompleted) { Stop(); return; }
            Replies.Tick();
        }
        private Task Send(PersonalReplyTicket ticket, CancellationToken cancellation)
        {
            return Task.Run(() => StaInputWorker.Instance.Invoke(() =>
            {
                Action require = () =>
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (!Replies.CanSend(ticket)) throw new OperationCanceledException();
                    if (!receiver.CanSendPersonal(ticket.Room))
                        throw new InvalidOperationException("대화 상대·방 이름·최신 수신 상태를 확인하지 못했습니다. 방 목록을 새로 읽고 다시 시작하세요.");
                };
                require();
                WindowsMacro.Instance.OpenChatRoom(ticket.Room.Name, ticket.Room.ProcessId);
                require();
                var input = new NativeMentionInput(ticket.Room.Name, ticket.Room.ProcessId);
                if (ticket.Action != "reply")
                {
                    PersonalActionSender.Send(ticket, input, require, () => Replies.Sending(ticket), () => Replies.ReactionSelected(ticket));
                    return;
                }
                Action requireText = () => { require(); PersonalActionSender.RequireNoDraft(ticket.Room); };
                InputMentionSnapshot owned = null;
                bool attempted = false;
                try
                {
                    string text = PersonalReplyCoordinator.Normalize(ChatProfanityFilter.Mask(ticket.Text)).Replace('\n', '\r');
                    if (string.IsNullOrWhiteSpace(text) || text.Length > 600 || char.IsControl(text[0])) throw new InvalidOperationException("답장 형식이 올바르지 않습니다.");
                    var empty = input.Read();
                    if (!empty.IsEmpty) throw new InvalidOperationException("작성 중인 메시지가 있어 자동응답을 멈췄습니다.");
                    requireText();
                    int seedLength = char.IsHighSurrogate(text[0]) ? 2 : 1;
                    input.BeginMention(empty, text.Substring(0, seedLength));
                    owned = input.Read();
                    if (owned.Text != text.Substring(0, seedLength) + "\r" || owned.Mentions.Count != 0)
                    { owned = null; throw new InvalidOperationException("답장 입력 중 내용이 바뀌었습니다."); }
                    requireText();
                    if (text.Length > seedLength) owned = input.Append(owned, text.Substring(seedLength));
                    if (owned.Text != text + "\r" || owned.Mentions.Count != 0) throw new InvalidOperationException("답장 입력 결과를 확인하지 못했습니다.");
                    requireText();
                    Replies.Sending(ticket);
                    attempted = true; // 이후 오류는 전송 여부가 불분명하므로 재시도/초안 삭제 금지
                    input.SendOnce(owned);
                    for (int i = 0; i < 20; i++) { Thread.Sleep(50); if (input.Read().IsEmpty) return; }
                    throw new InvalidOperationException("전송 결과를 확인하지 못했습니다. 자동 재전송하지 않습니다.");
                }
                finally
                {
                    if (!attempted && owned != null)
                    {
                        try { if (KakaoInputMentions.SameSnapshot(input.Read(), owned)) input.Clear(owned); }
                        catch (Exception) { /* 사용자 입력이 바뀌었으면 손대지 않습니다. */ }
                    }
                }
            }), cancellation);
        }
        internal Task Stop()
        {
            if (completion != null) return completion;
            var replies = Replies.Stop(); receiver.Dispose();
            catalogCancellation.Cancel();
            completion = Finish(replies);
            return completion;
        }
        private async Task Finish(Task replies)
        {
            try { await Task.WhenAll(replies, receiver.Completion, catalog).ConfigureAwait(false); }
            catch (Exception) { /* 취소/실패한 작업을 관찰하고 종료합니다. */ }
            finally { ai.Dispose(); catalogCancellation.Dispose(); }
        }
    }
}
