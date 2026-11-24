using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KakaotalkBot
{
    internal sealed class PersonalReplyPanel : UserControl
    {
        private PersonalReplySettings config;
        private PersonalReplyService service;
        private ChatCatalog catalog;
        private readonly Func<bool> canStart;
        private readonly Action reload;
        private readonly TextBox apiKey = new TextBox { Width = 180, UseSystemPasswordChar = true };
        private readonly TextBox model = new TextBox { Width = 155 };
        private readonly TextBox owner = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, MaxLength = 3000 };
        private readonly TextBox contact = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, MaxLength = 2000 };
        private readonly Label contactTitle = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, Text = "메모를 작성할 상대를 목록에서 선택하세요." };
        private readonly TextBox preview = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Dock = DockStyle.Fill };
        private readonly ListView people = new ListView { Dock = DockStyle.Fill, View = View.Details, CheckBoxes = true, FullRowSelect = true, HideSelection = false, MultiSelect = false };
        private readonly Label status = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, Text = "일반 1:1 대화를 선택한 뒤 미리보기로 확인하세요." };
        private readonly Label keyState = new Label { AutoSize = true, Padding = new Padding(0, 5, 0, 0) };
        private readonly CheckBox history = new CheckBox { Text = "최근 대화 80개로 말투·맥락 참고", AutoSize = true, Checked = true };
        private readonly Button save, test, refresh, draft, start, stop, resume;
        private ChatRoomInfo editing;
        private bool loading, bindingContact, connecting;
        private CancellationTokenSource connectionCancellation;
        private Task connectionTask;
        internal bool Busy { get { return service != null || connecting; } }

        internal PersonalReplyPanel(Func<bool> canStart, Action reload, PersonalReplySettings settings = null)
        {
            this.canStart = canStart; this.reload = reload;
            Dock = DockStyle.Fill;
            try { config = settings ?? PersonalReplySettings.Load(); }
            catch (Exception) { config = new PersonalReplySettings(); status.Text = "저장된 개인 AI 설정을 읽지 못했습니다. 설정을 다시 입력하세요."; }
            model.Text = config.Model; owner.Text = config.OwnerNotes; history.Checked = config.ImportHistory;
            keyState.Text = string.IsNullOrEmpty(config.ProtectedApiKey) ? "키 미등록" : "키 저장됨";
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 8, Padding = new Padding(8) };
            foreach (int height in new[] { 56, 34, 28, 70 }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "개인 AI 대리 응답\n선택한 상대와의 최근 대화·메모를 OpenAI로 보내 답장을 만듭니다. 직접 답장하면 해당 상대의 자동응답이 멈춥니다. 앱을 다시 켜면 중지 상태로 시작합니다." }, 0, 0);
            var api = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            api.Controls.AddRange(new Control[] { Caption("API 키"), apiKey, Caption("모델"), model, keyState });
            layout.Controls.Add(api, 0, 1);
            layout.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "내 말투·알려줘도 되는 정보 (현재 일정 등은 확인된 내용만 작성)" }, 0, 2);
            layout.Controls.Add(owner, 0, 3);
            var middle = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 49)); middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 51));
            people.Columns.Add("응답할 상대", 140); people.Columns.Add("상태", 235);
            middle.Controls.Add(people, 0, 0);
            var detail = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
            detail.RowStyles.Add(new RowStyle(SizeType.Absolute, 24)); detail.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
            detail.RowStyles.Add(new RowStyle(SizeType.Absolute, 24)); detail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            detail.Controls.Add(contactTitle, 0, 0);
            detail.Controls.Add(contact, 0, 1);
            detail.Controls.Add(new Label { Text = "최근 답장 미리보기 / 직접 확인할 내용", Dock = DockStyle.Fill }, 0, 2);
            detail.Controls.Add(preview, 0, 3); middle.Controls.Add(detail, 1, 0); layout.Controls.Add(middle, 0, 4);
            var options = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            options.Controls.AddRange(new Control[] { history, Caption("답장 길이에 따라 1~20초 대기 · 실행당 AI 요청 최대 100회") });
            layout.Controls.Add(options, 0, 5);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill };
            refresh = Button("대화 목록 새로고침"); save = Button("설정 저장"); test = Button("연결 확인");
            draft = Button("미리보기 시작"); start = Button("자동응답 시작"); stop = Button("전체 중지"); resume = Button("선택 상대 재개");
            actions.Controls.AddRange(new Control[] { refresh, save, test, draft, start, stop, resume });
            layout.Controls.Add(actions, 0, 6); layout.Controls.Add(status, 0, 7); Controls.Add(layout);
            people.ItemCheck += (s, e) =>
            {
                if (loading) return;
                if (Busy) { e.NewValue = e.CurrentValue; return; }
                if (e.NewValue == CheckState.Checked)
                {
                    // 체크박스 클릭은 행 선택을 바꾸지 않을 수 있으므로 메모 대상도 명시적으로 선택합니다.
                    var item = people.Items[e.Index];
                    item.Selected = true; item.Focused = true;
                    SelectContact(item.Tag as ChatRoomInfo);
                }
            };
            people.SelectedIndexChanged += (s, e) =>
            {
                if (loading || bindingContact) return;
                SelectContact(people.SelectedItems.Count == 0 ? null : people.SelectedItems[0].Tag as ChatRoomInfo);
            };
            contact.TextChanged += (s, e) => SaveContact();
            refresh.Click += (s, e) => { if (!Busy) reload(); };
            save.Click += (s, e) => TrySave();
            test.Click += async (s, e) => { connectionTask = TestConnection(); await connectionTask; };
            draft.Click += async (s, e) => await Start(false);
            start.Click += async (s, e) => await Start(true);
            stop.Click += async (s, e) => await Stop();
            resume.Click += (s, e) => { if (service != null && editing != null) service.Replies.Resume(editing.ChatId); };
            Pulse();
        }
        private static Label Caption(string text) { return new Label { Text = text, AutoSize = true, Padding = new Padding(0, 5, 0, 0) }; }
        private static Button Button(string text) { return new Button { Text = text, AutoSize = true, Height = 28 }; }
        private PersonalContactSettings ContactSettings(ChatRoomInfo room)
        {
            var value = config.Contacts.FirstOrDefault(c => c.AccountPath == room.AccountPath && c.ChatId == room.ChatId && c.UserId == room.DirectUserId);
            if (value == null) { value = new PersonalContactSettings { AccountPath = room.AccountPath, ChatId = room.ChatId, UserId = room.DirectUserId }; config.Contacts.Add(value); }
            return value;
        }
        private void SaveContact()
        {
            if (editing != null && !loading && !bindingContact && !Busy) ContactSettings(editing).Notes = contact.Text;
        }
        private void SelectContact(ChatRoomInfo room)
        {
            if (ReferenceEquals(editing, room))
            {
                contact.Enabled = !Busy && editing != null;
                if (editing == null) { contactTitle.Text = "메모를 작성할 상대를 목록에서 선택하세요."; preview.Clear(); }
                return;
            }
            SaveContact(); bindingContact = true;
            try
            {
                editing = room;
                contact.Text = editing == null ? "" : ContactSettings(editing).Notes;
                contactTitle.Text = editing == null ? "메모를 작성할 상대를 목록에서 선택하세요." : editing.Name + " · 관계·말투·주의할 내용";
                contact.Enabled = !Busy && editing != null;
            }
            finally { bindingContact = false; }
            UpdatePreview();
        }
        internal void SetCatalog(ChatCatalog value)
        {
            if (Busy) return;
            var previous = editing;
            SaveContact(); catalog = value; loading = true;
            try
            {
                people.Items.Clear(); editing = null; contact.Clear(); preview.Clear();
                if (value == null) return;
                foreach (var room in value.Rooms.Where(r => r.IsPersonalDirect).OrderBy(r => r.Name))
                    people.Items.Add(new ListViewItem(new[] { room.Name, "대기" }) { Tag = room });
                if (previous != null)
                {
                    var selected = people.Items.Cast<ListViewItem>().FirstOrDefault(item =>
                    {
                        var room = (ChatRoomInfo)item.Tag;
                        return room.AccountPath == previous.AccountPath && room.ChatId == previous.ChatId && room.DirectUserId == previous.DirectUserId;
                    });
                    if (selected != null) selected.Selected = true;
                }
                status.Text = "일반 1:1 대화 " + people.Items.Count + "개 · 같은 계정의 상대를 최대 10명 선택하세요.";
            }
            finally
            {
                loading = false;
                SelectContact(people.SelectedItems.Count == 0 ? null : people.SelectedItems[0].Tag as ChatRoomInfo);
            }
        }
        private bool TrySave()
        {
            try
            {
                SaveContact();
                if (!string.IsNullOrWhiteSpace(apiKey.Text)) config.SetApiKey(apiKey.Text);
                config.Model = model.Text.Trim(); config.OwnerNotes = owner.Text; config.ImportHistory = history.Checked;
                if (string.IsNullOrWhiteSpace(config.Model)) throw new InvalidOperationException("사용할 모델을 입력하세요.");
                config.Save(); apiKey.Clear(); keyState.Text = string.IsNullOrEmpty(config.ProtectedApiKey) ? "키 미등록" : "키 저장됨";
                status.Text = "설정 저장 완료"; return true;
            }
            catch (Exception error) { status.Text = error.Message; return false; }
        }
        private async Task TestConnection()
        {
            if (Busy || !TrySave()) return;
            connecting = true; connectionCancellation = new CancellationTokenSource(); Pulse();
            try
            {
                using (var ai = new PersonalReplyAi(config.GetApiKey(), config.Model))
                {
                    var result = await ai.Generate(new PersonalReplyPrompt { OwnerNotes = "친근한 반말", ContactNotes = "연결 테스트", History = new System.Collections.Generic.List<PersonalTurn> {
                        new PersonalTurn { Own = false, Text = "안녕!", At = DateTimeOffset.UtcNow.ToUnixTimeSeconds() } } }, connectionCancellation.Token);
                    preview.Text = result.Action == "reply" ? result.Text : result.Reason;
                    status.Text = "OpenAI 연결 확인 완료 · 카카오톡으로 보내지 않았습니다.";
                }
            }
            catch (Exception error) { status.Text = error is OperationCanceledException ? "연결 확인 시간 초과" : error.Message; }
            finally { connecting = false; connectionCancellation.Dispose(); connectionCancellation = null; Pulse(); }
        }
        private async Task Start(bool automatic)
        {
            if (Busy) return;
            try
            {
                if (!canStart()) throw new InvalidOperationException("기존 채팅 봇의 DB 수신을 중지하고 방 목록 읽기가 끝난 뒤 시작하세요.");
                var rooms = people.CheckedItems.Cast<ListViewItem>().Select(i => (ChatRoomInfo)i.Tag).ToArray();
                if (rooms.Length == 0 || rooms.Length > 10 || rooms.Any(r => r.AccountPath != rooms[0].AccountPath))
                    throw new InvalidOperationException("같은 계정의 일반 1:1 상대를 1~10명 선택하세요.");
                if (rooms.Any(r => !ChatDatabaseReceiver.MatchesPersonalRoom(r, catalog.Rooms)))
                    throw new InvalidOperationException("이름이 중복된 대화가 있습니다. 카카오톡에서 구분되는 방 이름을 지정하고 목록을 다시 읽어 주세요.");
                if (!TrySave()) return;
                service = new PersonalReplyService(rooms, config, automatic);
                status.Text = automatic ? "자동응답 시작 · 새로 오는 메시지에만 응답합니다." : "미리보기 시작 · 답장을 만들지만 전송하지 않습니다.";
            }
            catch (Exception error) { status.Text = error.Message; }
            Pulse(); await Task.CompletedTask;
        }
        internal async Task Stop()
        {
            if (connecting) { connectionCancellation?.Cancel(); if (connectionTask != null) await connectionTask; }
            var previous = service;
            if (previous == null) return;
            status.Text = "개인 AI 중지 중…";
            await previous.Stop();
            if (ReferenceEquals(service, previous)) service = null;
            status.Text = "개인 AI 중지됨 · 새 메시지에도 응답하지 않습니다.";
            Pulse();
        }
        private void UpdatePreview()
        {
            if (service == null || editing == null) { preview.Clear(); return; }
            var view = service.Replies.View().FirstOrDefault(v => v.ChatId == editing.ChatId);
            preview.Text = view == null ? "" : view.Preview;
        }
        internal void Pulse()
        {
            bool busy = Busy;
            apiKey.Enabled = model.Enabled = owner.Enabled = history.Enabled = save.Enabled = test.Enabled = refresh.Enabled = draft.Enabled = start.Enabled = !busy;
            contact.Enabled = !busy && editing != null;
            stop.Enabled = connecting || service != null && !service.IsStopping; resume.Enabled = service != null && !service.IsStopping && editing != null;
            if (service == null) return;
            if (service.IsStopping && service.Completion.IsCompleted)
            {
                string reason = service.Status; service = null; Pulse(); status.Text = "개인 AI 중지됨 · " + reason; return;
            }
            service.Tick();
            var views = service.Replies.View();
            foreach (ListViewItem item in people.Items)
            {
                var room = (ChatRoomInfo)item.Tag; var view = views.FirstOrDefault(v => v.ChatId == room.ChatId && room.AccountPath == ((ChatRoomInfo)people.CheckedItems[0].Tag).AccountPath);
                item.SubItems[1].Text = view == null ? "선택하지 않음" : view.Status;
            }
            status.Text = service.IsStopping ? "개인 AI 중지 중…" : (service.Automatic ? "자동응답 · " : "미리보기 (전송 없음) · ") + service.Status;
            UpdatePreview();
        }
    }
}
