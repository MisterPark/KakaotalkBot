using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using KakaotalkBot;
using Newtonsoft.Json.Linq;

internal static partial class PersonalReplyTests
{
    private static string ActionResponse(string action, string choice, string text = "")
    {
        var response = JObject.Parse(Response(action, text)); var content = response["output"][0]["content"][0];
        var decision = JObject.Parse((string)content["text"]); decision["choice"] = choice; content["text"] = decision.ToString(); return response.ToString();
    }
    private static ChatMessage Sticker(long chat, long log, string id, bool own, bool history)
    { var m = Message(chat, log, "[카카오톡 이모티콘 " + id + "]", own, history); m.EmoticonId = id; return m; }
    private static void ActionChoices()
    {
        DraftMemoryChecks();
        PickerReuseChecks();
        var item = new PersonalEmoticon { Pack = "2227359", Index = 1, Title = "테스트 팩", Meaning = "웃음" };
        var items = new[] { item };
        var reaction = PersonalReplyAi.Parse(ActionResponse("reaction", "1200509_021"));
        Check(PersonalReplyActions.Allowed(reaction, items, 25), "Known reaction requires a concrete incoming message");
        Check(!PersonalReplyActions.Allowed(reaction, items, 0), "No reaction without an incoming log id");
        var sticker = PersonalReplyAi.Parse(ActionResponse("emoticon", item.Id));
        Check(PersonalReplyActions.Allowed(sticker, items, 25), "Owned emoticon can be selected by stable id");
        Check(!PersonalReplyActions.Allowed(sticker, new PersonalEmoticon[0], 25), "Unowned or unavailable emoticon rejected");
        Check(!PersonalReplyActions.Allowed(new PersonalReplyDecision { Action = "reaction", Choice = "fake", Text = "" }, items, 25), "Unknown reaction rejected");
        Reject(() => PersonalReplyAi.Parse(ActionResponse("reaction", "1200509_021", "함께 보낼 텍스트")));
        Reject(() => PersonalReplyAi.Parse(ActionResponse("emoticon", "")));
        Reject(() => PersonalReplyAi.Parse(ActionResponse("reply", item.Id, "안녕")));
        var prompt = new PersonalReplyPrompt { History = new List<PersonalTurn>(), Emoticons = items.ToList(), ReactionTarget = 25 };
        var request = PersonalReplyAi.BuildRequest("test-model", prompt);
        var context = JObject.Parse((string)request["input"][0]["content"]);
        Check((string)context["available_emoticons"][0]["id"] == item.Id && context["available_reactions"].Count() == 6, "Model receives concrete permitted options");
        Check(request["text"]["format"]["schema"]["properties"]["action"]["enum"].Values<string>().Contains("emoticon"), "Structured output includes emoticon action");
        foreach (var type in new[] { 12, 20, 25 })
            Check(ChatMessageDecoder.EmoticonIdentity(type, "{\"path\":\"dw/2227359.emot_001.webp\"}") == item.Id, "Sticker formats retain stable identity " + type);
        Check(ChatMessageDecoder.EmoticonIdentity(2, "{\"path\":\"dw/2227359.emot_001.png\"}") == null, "Photo cannot impersonate an emoticon");
        Check(ChatMessageDecoder.EmoticonIdentity(12, "{\"path\":{}}") == null, "Malformed attachment does not crash receiver");
        Check(ChatMessageDecoder.EmoticonIdentity(12, "{\"path\":\"unrecognized.png\"}") == null, "Unknown paths are not invented ids");
        var labels = new Dictionary<long,string> { {1,"웃음"}, {2,"고마움"} };
        string emots = "[{\"itemCode\":\"2227359\",\"emotIdx\":1,\"soundPath\":\"\"},{\"itemCode\":\"2227359\",\"emotIdx\":2,\"soundPath\":\"audio\"}]";
        var row = new object[] { "2227359", "테스트 팩", "STICKER", 3, 99999L, emots, "[[1],[2]]" };
        var catalog = PersonalReplyActions.ReadCatalog(new[] {row,row}, labels, 10);
        Check(catalog.Count == 1 && catalog[0].Id == item.Id && catalog[0].Meaning == "웃음", "Catalog deduplicates and excludes sound items without losing tags");
        Check(PersonalReplyActions.ReadCatalog(new[] {row}, labels, 100000).Count == 0, "Expired items excluded");
        var malformed = (object[])row.Clone(); malformed[5] = "[{\"itemCode\":{},\"emotIdx\":999999999999999},{\"itemCode\":\"2227359\",\"emotIdx\":{}}]";
        Check(PersonalReplyActions.ReadCatalog(new[] {malformed,row}, labels, 10).Count == 1, "One malformed catalog entry does not discard valid choices");
        var many = Enumerable.Range(0, 1000).Select(n => new PersonalEmoticon { Pack = (1000000 + n / 10).ToString(), Index = n % 10 + 1, Title = "테스트 팩 " + n / 10, Meaning = n == 998 ? "감사·고마움" : "평소" }).ToList();
        var manyUsage = new JObject(new JProperty("owner_favorites", new JArray(new JObject(new JProperty("id", many[999].Id), new JProperty("count", 20)))));
        var candidates = PersonalReplyActions.Candidates(many, new[] { new PersonalTurn { Text = "정말 고마워", LogId = 91 } }, manyUsage);
        Check(candidates.Count == 240 && candidates.Select(e => e.Id).Distinct().Count() == 240, "Large catalog produces bounded unique AI choices");
        Check(candidates.Any(e => e.Id == many[999].Id) && candidates.Any(e => e.Id == many[998].Id), "Favorites and relevant meanings beyond the former 400-item cutoff remain selectable");
        Check(candidates.Select(e => e.Pack).Distinct().Count() == 100, "Many owned packs get candidate coverage instead of prefix truncation");
        now = new DateTime(2026, 11, 24, 0, 0, 0, DateTimeKind.Utc);
        var usage = new PersonalEmoticonHistory();
        usage.Add(Sticker(10, 1, "1111111:2", false, true)); now = now.AddSeconds(10); usage.Add(Sticker(10, 2, item.Id, true, true));
        usage.Add(Sticker(10, 3, "1111111:2", false, true)); usage.Add(Message(10, 4, "다른 화제")); usage.Add(Sticker(10, 5, item.Id, true, true));
        usage.Add(Sticker(10, 6, "1111111:2", false, true)); now = now.AddSeconds(901); usage.Add(Sticker(10, 7, item.Id, true, true));
        var history = usage.Snapshot(items);
        Check((int)history["owner_favorites"][0]["count"] == 3 && (int)history["contact_favorites"][0]["count"] == 3, "Own and contact frequencies stay separate");
        Check(history["reply_pairs"].Count() == 1 && (int)history["reply_pairs"][0]["count"] == 1, "Pairs exclude intervening messages and long time gaps");
        Check(usage.Snapshot(new PersonalEmoticon[0])["reply_pairs"].Count() == 0, "Past pairs do not authorize unavailable stickers");
        var otherUsage = new PersonalEmoticonHistory();
        Check(otherUsage.Snapshot(items)["owner_favorites"].Count() == 0, "Usage statistics are per conversation");
        for (int n = 0; n < 1010; n++) usage.Add(Message(10, 100+n, "새 대화"));
        Check((int)usage.Snapshot(items)["sample_messages"] == 1000 && usage.Snapshot(items)["owner_favorites"].Count() == 0, "Statistics have a bounded rolling window");

        var fake = new FakeAi(); var room = Room(10,110,"테스트 상대"); PersonalReplyTicket sent = null;
        var engine = new PersonalReplyCoordinator(new[] {room}, new PersonalReplySettings(), fake, true,
            (ticket, token) => { sent = ticket; return Task.FromResult(0); }, () => now, () => items);
        engine.Receive(Sticker(10, 50, "1111111:2", false, true)); engine.Receive(Sticker(10, 51, item.Id, true, true));
        engine.Receive(Sticker(10, 52, "1111111:2", false, false)); now = now.AddSeconds(1); engine.Tick(); Wait(() => fake.Count == 1);
        Check(fake.At(0).Prompt.ReactionTarget == 52 && fake.At(0).Prompt.History.Last().EmoticonId == "1111111:2", "Incoming sticker reaches AI with exact reaction target");
        Check(fake.At(0).Prompt.EmoticonHistory["reply_pairs"].Count() == 1, "Observed sticker response pattern reaches this contact's AI prompt");
        fake.At(0).Result.SetResult(sticker); Wait(() => engine.View()[0].Preview.StartsWith("[이모티콘]"), engine.Tick);
        now = now.AddSeconds(1); engine.Tick(); Wait(() => sent != null);
        Check(sent.Action == "emoticon" && sent.Choice == item.Id && sent.TargetLogId == 52, "Only the chosen sticker and original incoming log are queued");
        engine.Sending(sent); engine.Receive(Sticker(10, 53, item.Id, true, false)); engine.Tick();
        Check(engine.View()[0].Status.Contains("전송 확인"), "Sticker echo is not mistaken for manual takeover");
        engine.Receive(Sticker(10, 54, "2227359:9", true, false));
        Check(engine.View()[0].Status.Contains("직접 답장"), "Different manual sticker pauses automatic responses");
        engine.Stop().GetAwaiter().GetResult();
    }
    private sealed class ActionMemoryFixture
    {
        private readonly Dictionary<ulong, byte> data = new Dictionary<ulong, byte>();
        private ulong next = 0x100000;
        internal ulong Room, Vm, Floating, Controller, View, Model, Item;
        internal ulong Allocate(int size)
        { ulong at = next; next += (ulong)((size + 4095) / 4096) * 4096; Put(at, new byte[size]); return at; }
        internal void Put(ulong at, byte[] bytes) { for (int n = 0; n < bytes.Length; n++) data[at + (ulong)n] = bytes[n]; }
        internal void U(ulong at, ulong value) { Put(at, BitConverter.GetBytes(value)); }
        internal void I(ulong at, int value) { Put(at, BitConverter.GetBytes(value)); }
        internal byte[] Read(ulong at, int count)
        { var bytes = new byte[count]; for (int n = 0; n < count; n++) if (!data.TryGetValue(at + (ulong)n, out bytes[n])) return null; return bytes; }
        internal void Type(ulong at, string name)
        {
            ulong module = Allocate(4096), locator = module + 0x100, table = module + 0x300;
            U(table - 8, locator); I(locator, 1); I(locator + 12, 0x200); I(locator + 20, 0x100);
            Put(module + 0x210, System.Text.Encoding.ASCII.GetBytes(name + "\0")); U(at, table);
        }
        internal void Text(ulong field, string value)
        { ulong at = Allocate(128); Put(at, System.Text.Encoding.UTF8.GetBytes(value + "\0")); U(field, at); }
        internal ActionMemoryFixture()
        {
            Room = Allocate(4096); Vm = Allocate(4096); Floating = Allocate(4096);
            ulong shared = Allocate(512); Controller = shared + 16; View = Allocate(512); Model = Allocate(256); Item = Allocate(256);
            Type(Room, ".?AVCChatDlg@@"); Type(Vm, ".?AVChatRoom_ViewModel@@"); Type(Floating, ".?AVChatDelegate_Floating@@");
            Type(Floating + 0x580, ".?AVFloatingDecoView@@"); Type(shared, ".?AV?$_Ref_count_obj2@VChatLogEmoticonDecoController@@@std@@");
            Type(Controller, ".?AVChatLogEmoticonDecoController@@"); Type(View, ".?AVChatLogEmoticonDecoView@@");
            Type(Model, ".?AVEmoticonObjectModelImpl@@"); Type(Item, ".?AUEmoticonExItem@@");
            U(Room + 0x640, Vm); U(Vm + 0xE0, Room); U(Vm + 0xE8, 0xCAFE00); U(Vm + 0x2A8, Floating);
            U(Floating + 0x18, Room); U(Floating + 0x28, Vm); U(Floating + 0x38, Room + 0x650);
            U(Floating + 0x850, Room + 0x650); U(Floating + 0x870, Controller); U(Floating + 0x878, shared);
            U(Controller + 0x30, Room); U(Controller + 0x38, 0xCAFE00); U(Controller + 0x28, View);
            U(View + 0x18, Floating + 0x588); U(View + 0x28, Controller + 8); U(View + 0x40, Model);
            U(Controller + 0x50, Model); U(Controller + 0x58, Item); I(Item + 8, 4);
            Text(Item + 0x10, "4412206"); Text(Model + 0x40, "4412206"); Text(Model + 0x18, "4412206.emot_004.webp");
        }
        internal PersonalActionMemory Reader() { return new PersonalActionMemory(Read, (w, n) => IntPtr.Zero); }
    }
    private static void DraftMemoryChecks()
    {
        var f = new ActionMemoryFixture();
        using (var memory = f.Reader())
        {
            ulong instance;
            Check(memory.Draft(f.Room, out instance) == "4412206:4" && instance == f.Model, "Pending sticker identity comes from the chat-log preview controller");
            f.U(f.Controller + 0x50, 0); f.U(f.Controller + 0x58, 0);
            Check(memory.Draft(f.Room, out instance) == null && instance == 0, "Canceled preview's cached image is not a pending draft");
            f.U(f.Controller + 0x50, f.Model); f.U(f.Controller + 0x58, f.Item);
            f.Text(f.Model + 0x18, "4412206.emot_005.webp"); Reject(() => memory.Draft(f.Room));
            f.Text(f.Model + 0x18, "4412206.emot_004.webp"); f.Text(f.Model + 0x40, "2227359"); Reject(() => memory.Draft(f.Room));
            f.Text(f.Model + 0x40, "4412206"); f.U(f.Controller + 0x58, 0); Reject(() => memory.Draft(f.Room));
            f.U(f.Controller + 0x58, f.Item); f.U(f.Controller + 0x30, f.Room + 16); Reject(() => memory.Draft(f.Room));
            f.U(f.Controller + 0x30, f.Room); f.U(f.View + 0x40, f.Model + 16); Reject(() => memory.Draft(f.Room));
        }
    }
    private static void PickerReuseChecks()
    {
        var existing = new object(); int opens = 0, waits = 0, validations = 0;
        var reused = PersonalActionSender.OpenOrReusePicker(() => existing, () => opens++, () => { waits++; return new object(); }, p => { Check(ReferenceEquals(p, existing), "Existing picker is the one validated"); validations++; });
        Check(ReferenceEquals(reused, existing) && opens == 0 && waits == 0 && validations == 1, "A picker left open after sending is reused without toggling the opener");
        var opened = PersonalActionSender.OpenOrReusePicker(() => (object)null, () => opens++, () => { waits++; return existing; }, p => validations++);
        Check(ReferenceEquals(opened, existing) && opens == 1 && waits == 1 && validations == 2, "Closed picker is opened exactly once and validated");
        Reject(() => PersonalActionSender.OpenOrReusePicker(() => existing, () => opens++, () => existing, p => { throw new InvalidOperationException("Wrong room"); }));
        Check(opens == 1, "An unrelated existing picker is rejected without toggling or reopening it");
        var f = new ActionMemoryFixture();
        ulong popup = f.Allocate(4096), vm = f.Allocate(512); var roomWindow = new IntPtr(0x12345);
        f.Type(popup, ".?AVCEmoticonWnd@@"); f.Type(popup + 0x658, ".?AVCEmoticonSelectWnd@@"); f.Type(vm, ".?AVEmoticonViewModel@@");
        f.U(popup + 0x630, vm); f.U(vm + 0xE8, (ulong)roomWindow.ToInt64()); f.U(vm + 0xF0, (ulong)roomWindow.ToInt64());
        using (var memory = f.Reader())
        {
            memory.EmoticonPicker(popup, roomWindow); Check(true, "Reusable emoticon picker is bound to the intended chat window");
            f.U(vm + 0xF0, 0x22222); Reject(() => memory.EmoticonPicker(popup, roomWindow));
            f.U(vm + 0xF0, (ulong)roomWindow.ToInt64()); f.U(vm + 0xE8, 0x22222); Reject(() => memory.EmoticonPicker(popup, roomWindow));
            ulong reaction = f.Allocate(128), shared = f.Allocate(128), reactionPopup = f.Allocate(4096);
            f.Type(reactionPopup, ".?AVReactionPickerWnd@@"); f.U(f.Vm + 0x370, reaction); f.U(reaction + 0x60, shared);
            f.U(shared + 0x40, 10); f.U(reactionPopup + 0x638, 25); f.U(reactionPopup + 0x5E0, reaction); f.U(reactionPopup + 0x630, shared);
            memory.ReactionContext(reactionPopup, f.Room, 10, 25); Check(true, "Existing reaction picker matches the intended message");
            Reject(() => memory.ReactionContext(reactionPopup, f.Room, 10, 26));
            Reject(() => memory.ReactionContext(reactionPopup, f.Room, 11, 25));
        }
    }
}
