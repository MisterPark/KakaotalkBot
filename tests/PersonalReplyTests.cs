using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using KakaotalkBot;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

internal static partial class PersonalReplyTests
{
    private static int checks;
    private static DateTime now;
    private static void Check(bool value, string description) { if (!value) throw new Exception(description); checks++; }
    private static void Reject(Action action) { try { action(); } catch (InvalidOperationException) { checks++; return; } catch (ArgumentException) { checks++; return; } throw new Exception("Expected rejection"); }
    private static void Wait(Func<bool> ready, Action tick = null)
    {
        for (int i = 0; i < 300; i++) { if (tick != null) tick(); if (ready()) return; Thread.Sleep(10); }
        throw new Exception("Test timed out");
    }
    private static ChatRoomInfo Room(long chat, long user, string name)
    { return new ChatRoomInfo { ChatId = chat, DirectUserId = user, Name = name, Kind = "DirectChat", AccountPath = "synthetic-account", ProcessId = 99 }; }
    private static ChatMessage Message(long chat, long log, string text, bool own = false, bool history = false)
    { return new ChatMessage { ChatId = chat, LogId = log, AuthorId = own ? 1 : chat + 100, IsDirect = true, IsOwn = own, IsHistory = history, Message = text, SendAt = new DateTimeOffset(now).ToUnixTimeSeconds() }; }
    private static PersonalReplyDecision Reply(string text) { return new PersonalReplyDecision { Action = "reply", Text = text, Reason = "" }; }
    private static string Response(string action, string text)
    {
        return new JObject { { "status", "completed" }, { "output", new JArray(new JObject {
            { "type", "message" }, { "content", new JArray(new JObject { { "type", "output_text" },
                { "text", new JObject { { "action", action }, { "text", text }, { "reason", "확인 필요" } }.ToString(Formatting.None) } }) } }) } }.ToString();
    }
    private sealed class FakeAi : IPersonalReplyAi
    {
        internal sealed class Call { internal PersonalReplyPrompt Prompt; internal CancellationToken Token; internal TaskCompletionSource<PersonalReplyDecision> Result = new TaskCompletionSource<PersonalReplyDecision>(); }
        private readonly List<Call> calls = new List<Call>();
        internal int Count { get { lock (calls) return calls.Count; } }
        internal Call At(int index) { lock (calls) return calls[index]; }
        public Task<PersonalReplyDecision> Generate(PersonalReplyPrompt prompt, CancellationToken token)
        { lock (calls) { var call = new Call { Prompt = prompt, Token = token }; calls.Add(call); return call.Result.Task; } }
    }
    private sealed class Handler : HttpMessageHandler
    {
        internal string Url, Authorization, Json;
        internal HttpStatusCode Code = HttpStatusCode.OK;
        internal string ErrorBody = "secret-error-details";
        internal TimeSpan? RetryAfter;
        internal int Requests;
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            token.ThrowIfCancellationRequested(); Requests++; Url = request.RequestUri.ToString(); Authorization = request.Headers.Authorization.ToString();
            Json = await request.Content.ReadAsStringAsync();
            var result = new HttpResponseMessage(Code) { Content = new StringContent(Code == HttpStatusCode.OK ? Response("reply", "응 ㅎㅎ") : ErrorBody) };
            if (RetryAfter.HasValue) result.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(RetryAfter.Value);
            return result;
        }
    }
    [STAThread] private static int Main(string[] arguments)
    {
#if PERSONAL_AI_UI
        UiPreview(); return 0;
#else
        if (arguments.Contains("--ui")) { UiPreview(); return 0; }
        try { Run(); Console.WriteLine("PASS: " + checks + " personal AI checks (synthetic contacts, fake HTTP, no chat sends)"); return 0; }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
#endif
    }
    private static void Run()
    {
        ContactNoteSelection();
        ActionChoices();
        var config = new PersonalReplySettings();
        Reject(() => config.GetApiKey()); config.SetApiKey("test-only-secret");
        Check(config.GetApiKey() == "test-only-secret" && !JsonConvert.SerializeObject(config).Contains("test-only-secret"), "API key is encrypted at rest");
        Reject(() => config.SetApiKey("key\r\nheader"));
        var a = Room(10, 110, "친구 A"); var b = Room(20, 120, "직장 동료 B");
        Check(a.IsPersonalDirect && !new ChatRoomInfo { Kind = "OD", LinkId = 3, DirectUserId = 110 }.IsPersonalDirect, "Open direct chats excluded");
        Check(ChatDatabaseReceiver.MatchesPersonalRoom(a, new[] { a, b }), "Unique personal target matches");
        Reject(() => new ChatDatabaseReceiver(new[] { new ChatRoomInfo { Kind = "OD" } }, true));
        var duplicate = Room(30, 130, "친구 A");
        Check(!ChatDatabaseReceiver.MatchesPersonalRoom(a, new[] { a, duplicate }), "Duplicate room title blocked");
        var changed = Room(10, 999, "친구 A");
        Check(!ChatDatabaseReceiver.MatchesPersonalRoom(a, new[] { changed }), "Changed recipient blocked");
        changed = Room(10, 110, "이름 변경");
        Check(!ChatDatabaseReceiver.MatchesPersonalRoom(a, new[] { changed }), "Changed title blocked");
        var prompt = new PersonalReplyPrompt { OwnerNotes = "반말", ContactNotes = "오래된 친구", History = new List<PersonalTurn> {
            new PersonalTurn { Own = true, Text = "ㅋㅋ 그렇지", At = 1 }, new PersonalTurn { Own = false, Text = "안녕", At = 2 } } };
        JObject body = PersonalReplyAi.BuildRequest(config.Model, prompt);
        Check((bool)body["store"] == false && (string)body["text"]["format"]["type"] == "json_schema", "Stateless structured request");
        Check(!body.ContainsKey("tools") && !body.ContainsKey("previous_response_id"), "No tool access or cross-contact response state");
        Check(PersonalReplyAi.Parse(Response("reply", "안녕 ㅎㅎ")).Text == "안녕 ㅎㅎ", "Natural reply parsing");
        Check(PersonalReplyAi.Parse(Response("reply", "씨발")).Text == "**", "Existing profanity filter retained");
        Check(PersonalReplyAi.Parse(Response("hold", "")).Action == "hold", "Hold decision parsed");
        Reject(() => PersonalReplyAi.Parse(Response("reply", " ")));
        Reject(() => PersonalReplyAi.Parse(Response("reply", new string('가', 601))));
        Reject(() => PersonalReplyAi.Parse(Response("other", "hello")));
        Reject(() => PersonalReplyAi.Parse("{\"status\":\"incomplete\",\"output\":[]}"));
        Reject(() => PersonalReplyAi.Parse("not json"));
        Check(PersonalReplyAi.Parse("{\"status\":\"completed\",\"output\":[{\"type\":\"message\",\"content\":[{\"type\":\"refusal\"}]}]}").Action == "hold", "Refusal not sent");
        var handler = new Handler();
        using (var ai = new PersonalReplyAi("test-key", config.Model, handler))
        {
            Check(ai.Generate(prompt, CancellationToken.None).GetAwaiter().GetResult().Text == "응 ㅎㅎ", "HTTP reply round trip");
            Check(handler.Url == "https://api.openai.com/v1/responses" && handler.Authorization == "Bearer test-key", "Fixed official endpoint and auth");
            Check((bool)JObject.Parse(handler.Json)["store"] == false, "Storage disabled on wire");
            handler.Code = HttpStatusCode.Unauthorized;
            try { ai.Generate(prompt, CancellationToken.None).GetAwaiter().GetResult(); throw new Exception("Expected 401"); }
            catch (InvalidOperationException error) { Check(error.Message.Contains("401") && !error.Message.Contains("secret-error-details"), "HTTP error redacts response body"); }
            foreach (var pair in new[] {
                new[] { "credit_balance_exhausted", "크레딧이 소진" },
                new[] { "organization_spend_limit_exceeded", "조직의 API 지출 한도" },
                new[] { "project_spend_limit_exceeded", "프로젝트의 지출 한도" },
                new[] { "organization_usage_limit_exceeded", "승인된 API 사용 한도" },
                new[] { "insufficient_quota", "API 잔액 또는 사용 한도" } })
            {
                handler.Code = (HttpStatusCode)429;
                handler.ErrorBody = new JObject { { "error", new JObject { { "code", pair[0] }, { "type", "insufficient_quota" }, { "message", "secret-key-and-organization" } } } }.ToString();
                int previousRequests = handler.Requests;
                try { ai.Generate(prompt, CancellationToken.None).GetAwaiter().GetResult(); throw new Exception("Expected 429"); }
                catch (InvalidOperationException error) { Check(error.Message.Contains(pair[1]) && !error.Message.Contains("secret-key"), "Specific quota cause without private raw response"); }
                Check(handler.Requests == previousRequests + 1, "Quota failures are not retried");
            }
            handler.ErrorBody = "{\"error\":{\"type\":\"rate_limit_error\",\"code\":\"slow_down\"}}";
            handler.RetryAfter = TimeSpan.FromSeconds(37);
            try { ai.Generate(prompt, CancellationToken.None).GetAwaiter().GetResult(); throw new Exception("Expected rate limit"); }
            catch (InvalidOperationException error) { Check(error.Message.Contains("요청 속도") && error.Message.Contains("37초"), "Retry-After shown for rate limiting"); }
        }
        Check(PersonalReplyAi.DescribeError(429, "<html>secret</html>").Contains("상세 원인을 확인하지 못"), "Unknown 429 does not guess billing cause");
        Check(!PersonalReplyAi.DescribeError(429, "{\"error\":{\"code\":{\"secret\":\"key\"}}}").Contains("secret"), "Malformed error shape is safe");
        now = new DateTime(2026, 11, 24, 0, 0, 0, DateTimeKind.Utc);
        var fake = new FakeAi(); int sends = 0;
        var engine = new PersonalReplyCoordinator(new[] { a, b }, config, fake, false, (ticket, token) => { sends++; return Task.FromResult(0); }, () => now);
        engine.Receive(Message(10, 1, "A에게만 쓰는 말투 ㅋㅋ", true, true)); engine.Receive(Message(20, 1, "B의 비공개 내용", false, true));
        now = now.AddSeconds(20); engine.Tick(); Check(fake.Count == 0, "Imported history never triggers replies");
        var stale = Message(10, 99, "오래된 연락"); stale.SendAt -= 600; engine.Receive(stale); engine.Tick();
        Check(fake.Count == 0, "Old offline backlog does not trigger replies");
        engine.Receive(Message(10, 2, "첫 메시지")); engine.Receive(Message(10, 2, "중복"));
        now = now.AddMilliseconds(500); engine.Tick(); Check(fake.Count == 0, "Debounce before reply");
        engine.Receive(Message(10, 3, "두 번째 메시지")); now = now.AddMilliseconds(750); engine.Tick();
        Check(fake.Count == 0, "New inbound restarts the one-second batching window");
        now = now.AddMilliseconds(250); engine.Tick(); Wait(() => fake.Count == 1);
        Check(fake.At(0).Prompt.History.Count == 4 && fake.At(0).Prompt.History.All(t => !t.Text.Contains("B의")), "Batching, dedup and contact isolation");
        fake.At(0).Result.SetResult(Reply("응 뭐해 ㅎㅎ")); Wait(() => engine.View().First().Preview.Length > 0, engine.Tick);
        Check(sends == 0 && engine.View().First().Status.Contains("전송하지"), "Preview never sends"); engine.Stop().GetAwaiter().GetResult();

        fake = new FakeAi(); PersonalReplyTicket pendingTicket = null; var delivery = new TaskCompletionSource<int>();
        engine = new PersonalReplyCoordinator(new[] { a }, config, fake, true, (ticket, token) => { pendingTicket = ticket; sends++; return delivery.Task; }, () => now);
        engine.Receive(Message(10, 10, "안녕")); now = now.AddSeconds(9); engine.Tick(); Wait(() => fake.Count == 1);
        engine.Receive(Message(10, 11, "내가 직접 답할게", true));
        Check(fake.At(0).Token.IsCancellationRequested, "Manual reply cancels AI request");
        fake.At(0).Result.SetResult(Reply("낡은 답장")); Thread.Sleep(20); engine.Tick();
        engine.Receive(Message(10, 12, "추가 연락")); now = now.AddSeconds(30); engine.Tick();
        Check(sends == 0 && fake.Count == 1 && engine.View().First().Status.Contains("직접 답장"), "Manual takeover remains paused");
        engine.Resume(10); engine.Tick(); Check(fake.Count == 1, "Resume does not replay old messages");
        engine.Receive(Message(10, 13, "새 대화")); now = now.AddSeconds(9); engine.Tick(); Wait(() => fake.Count == 2);
        fake.At(1).Result.SetResult(Reply("좋아 ㅎㅎ")); Wait(() => engine.View().First().Preview == "좋아 ㅎㅎ", engine.Tick);
        now = now.AddSeconds(1); Wait(() => pendingTicket != null, engine.Tick);
        Check(engine.CanSend(pendingTicket), "Current reply ticket accepted");
        engine.Receive(Message(10, 14, "정정할게"));
        Check(!engine.CanSend(pendingTicket), "New inbound invalidates queued reply");
        delivery.SetException(new OperationCanceledException()); Thread.Sleep(20); engine.Tick(); now = now.AddSeconds(9); engine.Tick(); Wait(() => fake.Count == 3);
        Check(fake.At(2).Prompt.History.Last().Text == "정정할게", "Cancelled queued send regenerates from newest message");
        fake.At(2).Result.SetResult(new PersonalReplyDecision { Action = "hold", Text = "", Reason = "일정 확인 필요" });
        Wait(() => engine.View().First().Status.Contains("직접 확인 필요"), engine.Tick);
        int beforeHold = sends; engine.Receive(Message(10, 15, "그래")); now = now.AddSeconds(20); engine.Tick(); Check(sends == beforeHold && fake.Count == 3, "Hold pauses automatic responses");
        engine.Stop().GetAwaiter().GetResult(); Check(!engine.CanSend(pendingTicket), "Stop revokes outstanding send tickets");

        fake = new FakeAi(); pendingTicket = null;
        engine = new PersonalReplyCoordinator(new[] { a }, config, fake, true, (ticket, token) => { pendingTicket = ticket; return Task.FromResult(0); }, () => now);
        engine.Receive(Message(10, 30, "안녕")); now = now.AddSeconds(9); engine.Tick(); Wait(() => fake.Count == 1);
        fake.At(0).Result.SetResult(Reply("안녕\nㅎㅎ")); Wait(() => engine.View().First().Preview == "안녕\nㅎㅎ", engine.Tick);
        now = now.AddSeconds(1); Wait(() => pendingTicket != null, engine.Tick);
        engine.Sending(pendingTicket); engine.Receive(Message(10, 31, "안녕\r\nㅎㅎ", true)); engine.Tick();
        Check(engine.View().First().Status.Contains("전송 확인"), "Own AI echo is not manual takeover");
        var unsupported = Message(10, 32, null); unsupported.IsUnsupported = true; engine.Receive(unsupported);
        Check(engine.View().First().Status.Contains("직접 확인"), "Unsupported media pauses instead of guessing"); engine.Stop().GetAwaiter().GetResult();

        fake = new FakeAi(); pendingTicket = null;
        engine = new PersonalReplyCoordinator(new[] { a }, config, fake, true, (ticket, token) => { pendingTicket = ticket; return Task.FromResult(0); }, () => now);
        engine.Receive(Message(10, 40, "테스트")); now = now.AddSeconds(9); engine.Tick(); Wait(() => fake.Count == 1);
        fake.At(0).Result.SetResult(Reply("응")); Wait(() => engine.View().First().Preview == "응", engine.Tick);
        now = now.AddSeconds(1); Wait(() => pendingTicket != null, engine.Tick); engine.Sending(pendingTicket);
        now = now.AddSeconds(61); engine.Tick(); Check(engine.View().First().Status.Contains("전송 기록 미확인"), "Missing receipt pauses without repeat sends");
        engine.Stop().GetAwaiter().GetResult();
        ReplyTiming();
    }
    private static void ReplyTiming()
    {
        // A legacy fixed delay must not slow down the new length-based timing.
        var config = JsonConvert.DeserializeObject<PersonalReplySettings>("{\"DelaySeconds\":60,\"OwnerNotes\":\"반말\",\"Contacts\":[]}");
        var room = Room(10, 110, "친구 A");
        foreach (var sample in new[] { new[] { 2, 1000 }, new[] { 40, 5000 }, new[] { 80, 10000 }, new[] { 600, 20000 } })
        {
            var fake = new FakeAi(); PersonalReplyTicket sent = null; int sends = 0;
            var engine = new PersonalReplyCoordinator(new[] { room }, config, fake, true,
                (ticket, token) => { Interlocked.Increment(ref sends); sent = ticket; return Task.FromResult(0); }, () => now);
            engine.Receive(Message(10, 100, "안녕")); now = now.AddSeconds(1); engine.Tick(); Wait(() => fake.Count == 1);
            var reply = new string('가', sample[0]); fake.At(0).Result.SetResult(Reply(reply));
            Wait(() => engine.View().First().Preview == reply, engine.Tick);
            now = now.AddMilliseconds(sample[1] - 1); engine.Tick(); Thread.Sleep(20);
            Check(sent == null && sends == 0, "No early send for " + sample[0] + " characters");
            now = now.AddMilliseconds(1); Wait(() => sent != null, engine.Tick);
            Check(sent.Text == reply && engine.CanSend(sent) && sent.Expires == now.AddSeconds(30), "Length deadline dispatches with full send lifetime");
            engine.Tick(); engine.Tick(); Thread.Sleep(20); Check(sends == 1, "Repeated ticks do not duplicate a delayed send");
            engine.Stop().GetAwaiter().GetResult();
        }
        Check(PersonalReplyCoordinator.ReplyDelay(string.Concat(Enumerable.Repeat("\U0001F600", 16))) == TimeSpan.FromSeconds(2), "Emoji surrogate pairs count as visible characters");
        Check(PersonalReplyCoordinator.ReplyDelay(string.Concat(Enumerable.Repeat("e\u0301", 16))) == TimeSpan.FromSeconds(2), "Combining marks do not double typing time");

        foreach (var interruption in new[] { "incoming", "manual", "stop", "stale" })
        {
            var fake = new FakeAi(); PersonalReplyTicket sent = null; int sends = 0;
            var engine = new PersonalReplyCoordinator(new[] { room }, config, fake, true,
                (ticket, token) => { Interlocked.Increment(ref sends); sent = ticket; return Task.FromResult(0); }, () => now);
            engine.Receive(Message(10, 200, "긴 답장 질문")); now = now.AddSeconds(1); engine.Tick(); Wait(() => fake.Count == 1);
            var reply = new string('가', 160); fake.At(0).Result.SetResult(Reply(reply));
            Wait(() => engine.View().First().Preview == reply, engine.Tick); now = now.AddSeconds(2);
            if (interruption == "incoming")
            {
                engine.Receive(Message(10, 201, "정정한 새 질문")); now = now.AddSeconds(1); engine.Tick(); Wait(() => fake.Count == 2);
                Check(fake.At(1).Prompt.History.Last().Text == "정정한 새 질문", "New inbound replaces a reply waiting for its length deadline");
                fake.At(1).Result.SetResult(Reply("새 답장")); Wait(() => engine.View().First().Preview == "새 답장", engine.Tick);
                now = now.AddSeconds(1); Wait(() => sent != null, engine.Tick);
                now = now.AddSeconds(30); engine.Tick(); Thread.Sleep(20);
                Check(sends == 1 && sent.Text == "새 답장", "Only the replacement is sent after the old deadline passes");
            }
            else
            {
                if (interruption == "manual") engine.Receive(Message(10, 201, "내가 직접 답장", true));
                if (interruption == "stop") engine.Stop().GetAwaiter().GetResult();
                now = now.AddSeconds(interruption == "stale" ? 301 : 30); engine.Tick(); Thread.Sleep(20);
                Check(sends == 0 && sent == null, interruption + " cancels a delayed reply without sending it");
            }
            engine.Stop().GetAwaiter().GetResult();
        }

        var quickAi = new FakeAi(); PersonalReplyTicket quickSent = null;
        var quick = new PersonalReplyCoordinator(new[] { room }, config, quickAi, true,
            (ticket, token) => { quickSent = ticket; return Task.FromResult(0); }, () => now);
        quick.Receive(Message(10, 300, "안녕")); now = now.AddSeconds(1); quick.Tick(); Wait(() => quickAi.Count == 1);
        quickAi.At(0).Result.SetResult(Reply("응")); Wait(() => quick.View().First().Preview == "응", quick.Tick);
        now = now.AddSeconds(1); Wait(() => quickSent != null, quick.Tick); quick.Sending(quickSent);
        var firstSentAt = now; quick.Receive(Message(10, 301, "응", true)); quick.Tick();
        quick.Receive(Message(10, 302, "다음 질문")); now = now.AddSeconds(1); quick.Tick(); Wait(() => quickAi.Count == 2);
        Check(now - firstSentAt < TimeSpan.FromSeconds(15), "Short follow-up is not blocked by the removed fifteen-second cooldown");
        quickSent = null; quickAi.At(1).Result.SetResult(Reply("그래")); Wait(() => quick.View().First().Preview == "그래", quick.Tick);
        now = now.AddSeconds(1); Wait(() => quickSent != null, quick.Tick);
        Check(quickSent.Text == "그래" && now - firstSentAt == TimeSpan.FromSeconds(2), "Short follow-up sends after batching and its own short delay");
        quick.Stop().GetAwaiter().GetResult();
    }
    private static T PanelField<T>(PersonalReplyPanel panel, string name)
    {
        return (T)typeof(PersonalReplyPanel).GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(panel);
    }
    private static void ContactNoteSelection()
    {
        var a = Room(10, 110, "친구 A"); var b = Room(20, 120, "직장 동료 B");
        var config = new PersonalReplySettings();
        config.Contacts.Add(new PersonalContactSettings { AccountPath = b.AccountPath, ChatId = b.ChatId, UserId = b.DirectUserId, Notes = "B에게는 존댓말" });
        var catalog = new ChatCatalog(); catalog.Rooms.Add(a); catalog.Rooms.Add(b);
        using (var panel = new PersonalReplyPanel(() => false, () => { }, config))
        {
            var list = PanelField<ListView>(panel, "people"); var editor = PanelField<TextBox>(panel, "contact");
            var handle = list.Handle; panel.SetCatalog(catalog);
            var first = list.Items.Cast<ListViewItem>().Single(i => ((ChatRoomInfo)i.Tag).ChatId == a.ChatId);
            var second = list.Items.Cast<ListViewItem>().Single(i => ((ChatRoomInfo)i.Tag).ChatId == b.ChatId);
            first.Checked = true; Application.DoEvents();
            Check(first.Selected && editor.Enabled, "Checking a recipient binds its note editor instead of accepting unbound text");
            editor.Text = "내 이름은 민수. A에게는 반말";
            first.Selected = false; second.Selected = true; Application.DoEvents();
            Check(editor.Text == "B에게는 존댓말", "Selecting B restores B's saved notes");
            editor.Text = "B에게만 공개하는 정보";
            second.Selected = false; first.Selected = true; Application.DoEvents();
            Check(editor.Text == "내 이름은 민수. A에게는 반말", "A's draft survives the A to B to A selection sequence");
            first.Checked = false; Application.DoEvents();
            Check(editor.Text == "내 이름은 민수. A에게는 반말", "Unchecking a reply target does not erase its notes");
            panel.SetCatalog(catalog); Application.DoEvents();
            Check(editor.Text == "내 이름은 민수. A에게는 반말", "Catalog refresh preserves the selected contact and draft");
            var restored = JsonConvert.DeserializeObject<PersonalReplySettings>(JsonConvert.SerializeObject(config));
            Check(restored.Contacts.Single(c => c.ChatId == a.ChatId).Notes == "내 이름은 민수. A에게는 반말" &&
                restored.Contacts.Single(c => c.ChatId == b.ChatId).Notes == "B에게만 공개하는 정보", "Distinct recipient notes survive settings serialization");
            second = list.Items.Cast<ListViewItem>().Single(i => ((ChatRoomInfo)i.Tag).ChatId == b.ChatId);
            second.Checked = true; Application.DoEvents();
            Check(second.Selected && editor.Text == "B에게만 공개하는 정보" && PanelField<Label>(panel, "contactTitle").Text.Contains(b.Name), "Checking another target displays that person's notes and name");
            panel.SetCatalog(new ChatCatalog()); Application.DoEvents();
            Check(!editor.Enabled && editor.Text == "" && !PanelField<Label>(panel, "contactTitle").Text.Contains(b.Name), "An empty recipient list does not accept notes without an owner");

            now = new DateTime(2026, 11, 24, 0, 0, 0, DateTimeKind.Utc);
            var fake = new FakeAi(); var engine = new PersonalReplyCoordinator(new[] { a, b }, restored, fake, false,
                (ticket, token) => { throw new Exception("Selection test must not send messages"); }, () => now);
            engine.Receive(Message(a.ChatId, 400, "너 이름 뭐야?")); now = now.AddSeconds(1); engine.Tick(); Wait(() => fake.Count == 1);
            Check(fake.At(0).Prompt.ContactNotes == "내 이름은 민수. A에게는 반말", "A's edited note reaches AI prompt creation");
            engine.Receive(Message(b.ChatId, 401, "안녕하세요")); now = now.AddSeconds(1); engine.Tick(); Wait(() => fake.Count == 2);
            var body = PersonalReplyAi.BuildRequest(restored.Model, fake.At(1).Prompt);
            var sentContext = JObject.Parse((string)body["input"][0]["content"]);
            Check((string)sentContext["contact_notes"] == "B에게만 공개하는 정보" && !sentContext.ToString().Contains("내 이름은 민수"), "B's API input includes B's note without leaking A's note");
            fake.At(0).Result.SetResult(Reply("민수야")); fake.At(1).Result.SetResult(Reply("안녕하세요"));
            engine.Stop().GetAwaiter().GetResult();
        }
    }
    private static void UiPreview()
    {
        Application.EnableVisualStyles();
        var form = new Form { Text = "Personal AI UI verification (synthetic data)", Width = 850, Height = 760 };
        var panel = new PersonalReplyPanel(() => false, () => { }, new PersonalReplySettings());
        var catalog = new ChatCatalog(); catalog.Rooms.Add(Room(10, 110, "친구 A")); catalog.Rooms.Add(Room(20, 120, "직장 동료 B"));
        form.Controls.Add(panel); panel.SetCatalog(catalog); Application.Run(form);
    }
}
