using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace KakaotalkBot
{
    internal sealed class PersonalReplyTicket
    {
        internal ChatRoomInfo Room;
        internal string Text;
        internal string Action = "reply", Choice;
        internal long TargetLogId;
        internal PersonalEmoticon Emoticon;
        internal long Version;
        internal DateTime Expires;
    }
    internal sealed class PersonalReplyView
    {
        internal long ChatId;
        internal string Name, Status, Preview;
    }
    internal sealed class PersonalReplyCoordinator
    {
        private sealed class Conversation
        {
            internal ChatRoomInfo Room;
            internal string Notes, Status = "새 메시지 대기", Preview = "", Echo;
            internal readonly List<PersonalTurn> History = new List<PersonalTurn>();
            internal readonly PersonalEmoticonHistory EmoticonHistory = new PersonalEmoticonHistory();
            internal readonly HashSet<long> Seen = new HashSet<long>();
            internal readonly Queue<long> SeenOrder = new Queue<long>();
            internal bool Paused, Pending;
            internal long Version;
            internal long LatestInboundAt, LatestInboundLogId;
            internal string EchoEmoticon;
            internal List<PersonalEmoticon> Options = new List<PersonalEmoticon>();
            internal DateTime Due, EchoExpires, SendDue;
            internal PersonalReplyTicket Ready;
            internal CancellationTokenSource RequestCancellation;
            internal Task<PersonalReplyDecision> Request;
            internal long RequestVersion;
            internal Task Sending;
            internal long SendingVersion;
        }
        private readonly object gate = new object();
        private readonly Dictionary<long, Conversation> conversations;
        private readonly IPersonalReplyAi ai;
        private readonly Func<PersonalReplyTicket, CancellationToken, Task> send;
        private readonly Func<DateTime> clock;
        private readonly CancellationTokenSource stopped = new CancellationTokenSource();
        private readonly string notes;
        private readonly bool automatic;
        private readonly Func<IReadOnlyList<PersonalEmoticon>> emoticons;
        private bool active = true;
        private int requests;
        internal PersonalReplyCoordinator(IEnumerable<ChatRoomInfo> rooms, PersonalReplySettings config, IPersonalReplyAi ai,
            bool automatic, Func<PersonalReplyTicket, CancellationToken, Task> send, Func<DateTime> clock = null, Func<IReadOnlyList<PersonalEmoticon>> emoticons = null)
        {
            this.ai = ai; this.send = send; this.automatic = automatic; this.clock = clock ?? (() => DateTime.UtcNow);
            notes = config.OwnerNotes;
            this.emoticons = emoticons ?? (() => new PersonalEmoticon[0]);
            conversations = rooms.ToDictionary(r => r.ChatId, r => new Conversation { Room = r,
                Notes = config.Contacts.FirstOrDefault(c => c.ChatId == r.ChatId && c.AccountPath == r.AccountPath && c.UserId == r.DirectUserId)?.Notes ?? "" });
        }
        internal void Receive(ChatMessage message)
        {
            lock (gate)
            {
                Conversation conversation;
                if (!active || !message.IsDirect || message.LogId <= 0 || !conversations.TryGetValue(message.ChatId, out conversation) ||
                    !message.IsOwn && message.AuthorId != conversation.Room.DirectUserId || !conversation.Seen.Add(message.LogId)) return;
                conversation.SeenOrder.Enqueue(message.LogId);
                while (conversation.SeenOrder.Count > 2048) conversation.Seen.Remove(conversation.SeenOrder.Dequeue());
                conversation.EmoticonHistory.Add(message);
                if (!string.IsNullOrWhiteSpace(message.Message))
                {
                    conversation.History.Add(new PersonalTurn { Own = message.IsOwn, Text = PersonalReplyAi.Limit(message.Message, 2000), At = message.SendAt, LogId = message.LogId, EmoticonId = message.EmoticonId });
                    while (conversation.History.Count > 80 || conversation.History.Sum(t => t.Text.Length) > 24000) conversation.History.RemoveAt(0);
                }
                if (message.IsHistory) return;
                long now = new DateTimeOffset(clock()).ToUnixTimeSeconds();
                if (!message.IsOwn && (message.SendAt <= 0 || message.SendAt > now + 60 || now - message.SendAt > 300)) return;
                if (message.IsOwn && clock() <= conversation.EchoExpires &&
                    (conversation.Echo != null && !message.IsUnsupported && Normalize(message.Message) == Normalize(conversation.Echo) ||
                     conversation.EchoEmoticon != null && message.EmoticonId == conversation.EchoEmoticon))
                {
                    conversation.Echo = null; conversation.EchoEmoticon = null;
                    conversation.Status = "답장 전송 확인 · 새 메시지 대기";
                    return;
                }
                conversation.Version++;
                conversation.RequestCancellation?.Cancel();
                conversation.Ready = null;
                if (message.IsOwn) { Pause(conversation, "직접 답장 감지 · 자동응답 중지"); return; }
                if (message.IsUnsupported) { Pause(conversation, "사진·음성 등은 직접 확인해 주세요."); return; }
                if (string.IsNullOrWhiteSpace(message.Message) || conversation.Paused) return;
                conversation.Pending = true;
                conversation.LatestInboundAt = message.SendAt;
                conversation.LatestInboundLogId = message.LogId;
                conversation.Due = clock().AddSeconds(1);
                conversation.Status = "연속 메시지 모으는 중";
            }
        }
        internal static string Normalize(string value) { return (value ?? "").Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n'); }
        internal static TimeSpan ReplyDelay(string text)
        {
            int characters = new StringInfo(Normalize(text)).LengthInTextElements;
            return TimeSpan.FromSeconds(Math.Max(1, Math.Min(20, characters / 8.0)));
        }
        private static void Pause(Conversation conversation, string reason)
        {
            conversation.Paused = true; conversation.Pending = false; conversation.Ready = null; conversation.Status = reason;
            conversation.Version++; conversation.RequestCancellation?.Cancel();
        }
        internal void Resume(long chat)
        {
            lock (gate)
            {
                Conversation conversation;
                if (!active || !conversations.TryGetValue(chat, out conversation) || conversation.Sending != null && !conversation.Sending.IsCompleted) return;
                conversation.Version++; conversation.RequestCancellation?.Cancel();
                conversation.Paused = false; conversation.Pending = false; conversation.Ready = null; conversation.Echo = null; conversation.EchoEmoticon = null;
                conversation.Status = "재개 · 다음 새 메시지부터 응답";
            }
        }
        internal void Tick()
        {
            lock (gate)
            {
                if (!active) return;
                foreach (var conversation in conversations.Values)
                {
                    if ((conversation.Pending || conversation.Ready != null) && new DateTimeOffset(clock()).ToUnixTimeSeconds() - conversation.LatestInboundAt > 300)
                    {
                        conversation.Pending = false; conversation.Ready = null; conversation.Version++; conversation.RequestCancellation?.Cancel();
                        conversation.Status = "오래된 메시지 응답 생략 · 새 메시지 대기";
                    }
                    if (conversation.Sending != null && conversation.Sending.IsCompleted)
                    {
                        if (conversation.Sending.IsFaulted)
                        {
                            var error = conversation.Sending.Exception.GetBaseException();
                            if (!(error is OperationCanceledException) || conversation.SendingVersion == conversation.Version)
                                Pause(conversation, "전송 보류: " + error.Message);
                        }
                        else if (conversation.Sending.IsCanceled && conversation.SendingVersion == conversation.Version)
                            Pause(conversation, "전송 취소 · 직접 확인 후 재개하세요.");
                        conversation.Sending = null;
                    }
                    if ((conversation.Echo != null || conversation.EchoEmoticon != null) && clock() > conversation.EchoExpires)
                    { conversation.Echo = null; conversation.EchoEmoticon = null; Pause(conversation, "전송 기록 미확인 · 자동 재전송하지 않습니다."); }
                    if (conversation.Request != null && conversation.Request.IsCompleted)
                    {
                        var completed = conversation.Request; conversation.Request = null;
                        conversation.RequestCancellation.Dispose(); conversation.RequestCancellation = null;
                        if (conversation.RequestVersion != conversation.Version || conversation.Paused)
                        { if (completed.IsFaulted) { var ignored = completed.Exception; } continue; }
                        if (completed.IsCanceled) { Pause(conversation, "AI 요청 취소/시간 초과 · 재개 후 새 메시지부터 응답"); continue; }
                        if (completed.IsFaulted) { Pause(conversation, completed.Exception.GetBaseException().Message); continue; }
                        var result = completed.Result;
                        conversation.Pending = false;
                        if (!PersonalReplyActions.Allowed(result, conversation.Options, conversation.LatestInboundLogId)) { Pause(conversation, "AI 응답 항목을 확인하지 못했습니다."); continue; }
                        conversation.Preview = PersonalReplyActions.Preview(result, conversation.Options);
                        if (result.Action == "hold") { Pause(conversation, "직접 확인 필요: " + result.Reason); continue; }
                        if (!automatic) { conversation.Status = "답장 미리보기 완료 · 전송하지 않음"; continue; }
                        conversation.Ready = new PersonalReplyTicket { Room = conversation.Room, Version = conversation.Version, Text = result.Text,
                            Action = result.Action, Choice = result.Choice, TargetLogId = conversation.LatestInboundLogId,
                            Emoticon = conversation.Options.FirstOrDefault(e => e.Id == result.Choice) };
                        conversation.SendDue = clock().Add(result.Action == "reply" ? ReplyDelay(result.Text) : TimeSpan.FromSeconds(1));
                    }
                    if (conversation.Ready != null)
                    {
                        if (conversation.Paused || conversation.Ready.Version != conversation.Version) conversation.Ready = null;
                        else if (clock() >= conversation.SendDue)
                        {
                            var ticket = conversation.Ready; conversation.Ready = null;
                            // 전송 작업의 유효 시간은 글자 수에 따른 대기가 끝난 시점부터 계산합니다.
                            ticket.Expires = clock().AddSeconds(30);
                            conversation.Status = "답장 전송 대기";
                            conversation.SendingVersion = conversation.Version;
                            conversation.Sending = Task.Run(() => send(ticket, stopped.Token), stopped.Token);
                        }
                        else conversation.Status = "답장 길이 대기 · " + Math.Ceiling((conversation.SendDue - clock()).TotalSeconds) + "초 남음";
                    }
                    if (!conversation.Paused && conversation.Pending && conversation.Request == null && conversation.Ready == null && conversation.Sending == null &&
                        conversation.Echo == null && conversation.EchoEmoticon == null && clock() >= conversation.Due &&
                        conversations.Values.Count(c => c.Request != null) < 2)
                    {
                        if (requests >= 100) { Pause(conversation, "이번 실행의 AI 요청 100회 한도 · 중지 후 다시 시작하세요."); continue; }
                        requests++;
                        conversation.RequestVersion = conversation.Version;
                        conversation.RequestCancellation = CancellationTokenSource.CreateLinkedTokenSource(stopped.Token);
                        conversation.Status = "AI 답장 작성 중";
                        var available = emoticons();
                        conversation.Options = PersonalReplyActions.Candidates(available, conversation.History, conversation.EmoticonHistory.Snapshot(available));
                        var prompt = new PersonalReplyPrompt { OwnerNotes = notes, ContactNotes = conversation.Notes,
                            Emoticons = conversation.Options.ToList(), ReactionTarget = conversation.LatestInboundLogId,
                            EmoticonHistory = conversation.EmoticonHistory.Snapshot(conversation.Options),
                            History = conversation.History.Select(t => new PersonalTurn { Own = t.Own, Text = t.Text, At = t.At, LogId = t.LogId, EmoticonId = t.EmoticonId }).ToList() };
                        var token = conversation.RequestCancellation.Token;
                        conversation.Request = Task.Run(() => ai.Generate(prompt, token), token);
                    }
                }
            }
        }
        internal bool CanSend(PersonalReplyTicket ticket)
        {
            lock (gate)
            {
                Conversation conversation;
                return active && automatic && clock() <= ticket.Expires && conversations.TryGetValue(ticket.Room.ChatId, out conversation) &&
                    !conversation.Paused && conversation.Version == ticket.Version && conversation.Room.Identity == ticket.Room.Identity &&
                    (ticket.Action == "reply" || ticket.TargetLogId == conversation.LatestInboundLogId && ticket.TargetLogId > 0);
            }
        }
        internal void Sending(PersonalReplyTicket ticket)
        {
            lock (gate)
            {
                if (!CanSend(ticket)) throw new OperationCanceledException("새 메시지 또는 직접 답장으로 이전 답장을 취소했습니다.");
                var conversation = conversations[ticket.Room.ChatId];
                conversation.Echo = ticket.Action == "reply" ? ticket.Text : null;
                conversation.EchoEmoticon = ticket.Action == "emoticon" ? ticket.Choice : null;
                conversation.EchoExpires = clock().AddSeconds(60);
                conversation.Status = "전송 결과 확인 중";
            }
        }
        internal void ReactionSelected(PersonalReplyTicket ticket)
        {
            lock (gate)
            {
                Conversation conversation;
                if (ticket.Action == "reaction" && conversations.TryGetValue(ticket.Room.ChatId, out conversation) && conversation.Version == ticket.Version)
                    conversation.Status = "공감 선택 완료 · 새 메시지 대기";
            }
        }
        internal List<PersonalReplyView> View()
        {
            lock (gate) return conversations.Values.Select(c => new PersonalReplyView { ChatId = c.Room.ChatId,
                Name = c.Room.Name, Status = c.Status, Preview = c.Preview }).ToList();
        }
        internal Task Stop()
        {
            lock (gate)
            {
                active = false; stopped.Cancel();
                foreach (var conversation in conversations.Values) { conversation.Ready = null; conversation.Version++; conversation.RequestCancellation?.Cancel(); }
                return Task.WhenAll(conversations.Values.SelectMany(c => new Task[] { c.Request, c.Sending }).Where(t => t != null));
            }
        }
    }
}
