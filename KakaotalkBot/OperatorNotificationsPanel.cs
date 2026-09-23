using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace KakaotalkBot
{
    public partial class Form1
    {
        private OperatorNotificationQueue operatorNotices;
        private ComboBox noticeRoom;
        private CheckBox noticeEnabled, noticeReentry, noticeMemo, noticeHistory, noticeMonthly, noticeReset, noticeMemoBody;
        private Label noticeStatus;
        private ListBox noticeLog;
        private TextBox noticePreview;
        private string noticeLogSignature;
        private sealed class NoticeRoomChoice
        {
            internal ChatRoomInfo Room;
            public override string ToString() { return Room.Name + " · " + Room.ChatId + " · " + Path.GetFileName(Room.AccountPath); }
        }
        private sealed class NoticeLogChoice
        {
            internal OperatorNotice Notice;
            public override string ToString()
            {
                string state = Notice.State == "pending" ? "대기" : Notice.State == "blocked" ? "전송 전 차단" : Notice.State == "cancelled" ? "취소" :
                    Notice.State == "submitted" ? "전송 입력 완료" : Notice.State == "attempted" ? "시도 중" : "결과 미확인";
                return OperationsStore.LocalTime(Notice.At) + " [" + state + "] " + (Notice.Error ?? Notice.Key);
            }
        }
        private void InitializeOperatorNotifications()
        {
            var page = new TabPage("운영진 알림방"); operationsViews.TabPages.Add(page);
            var config = Settings.Instance;
            noticeEnabled = new CheckBox { Text = "운영진 알림 전송 사용", Checked = config.OperatorAlertsEnabled, AutoSize = true };
            noticeRoom = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 720 };
            noticeReentry = new CheckBox { Text = "재입장", Checked = config.AlertReentry, AutoSize = true };
            noticeMemo = new CheckBox { Text = "메모 등록", Checked = config.AlertMemo, AutoSize = true };
            noticeHistory = new CheckBox { Text = "/이력 조회 결과", Checked = config.AlertHistory, AutoSize = true };
            noticeMonthly = new CheckBox { Text = "월간 집계", Checked = config.AlertMonthly, AutoSize = true };
            noticeReset = new CheckBox { Text = "포인트 초기화 결과", Checked = config.AlertReset, AutoSize = true };
            noticeMemoBody = new CheckBox { Text = "메모 본문 포함 (알림방에 공개됨)", Checked = config.AlertMemoBody, AutoSize = true };
            var controls = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 200, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, AutoScroll = true };
            var reload = new Button { Text = "방 목록 불러오기", AutoSize = true };
            var save = new Button { Text = "설정 저장", AutoSize = true };
            var test = new Button { Text = "테스트 알림 전송", AutoSize = true };
            var retry = new Button { Text = "전송 전 차단 항목 재시도", AutoSize = true };
            controls.Controls.AddRange(new Control[] { noticeEnabled, reload, noticeRoom, noticeReentry, noticeMemo, noticeHistory, noticeMonthly, noticeReset, noticeMemoBody, save, test, retry });
            noticeStatus = new Label { Dock = DockStyle.Bottom, Height = 50, Text = "방 목록을 불러오고 운영진만 참여한 방을 지정하세요. 해당 채팅창을 열어 두세요." };
            noticeLog = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            noticePreview = ReadBox(); noticePreview.Dock = DockStyle.Bottom; noticePreview.Height = 170;
            page.Controls.Add(noticeLog); page.Controls.Add(noticePreview); page.Controls.Add(controls); page.Controls.Add(noticeStatus);
            noticeLog.SelectedIndexChanged += (s, e) => { var item = noticeLog.SelectedItem as NoticeLogChoice; noticePreview.Text = item == null ? "" : item.Notice.Body.Replace("\r\n", "\n").Replace("\n", "\r\n"); };
            try { operatorNotices = new OperatorNotificationQueue(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OperatorNotifications", "outbox.json")); }
            catch (Exception error) { noticeStatus.Text = "알림 기록 읽기 실패: " + error.GetBaseException().Message; }
            reload.Click += (s, e) => LoadNoticeRooms();
            save.Click += (s, e) => SaveNoticeSettings();
            test.Click += (s, e) =>
            {
                try
                {
                    var choice = noticeRoom.SelectedItem as NoticeRoomChoice;
                    if (!noticeEnabled.Checked || !config.OperatorAlertsEnabled || choice == null || choice.Room.ChatId != config.OperatorAlertRoomId || choice.Room.AccountPath != config.OperatorAlertAccount)
                        throw new InvalidOperationException("알림방과 전송 사용 설정을 먼저 저장하세요.");
                    QueueNotice("test:" + Guid.NewGuid().ToString("N"), "[운영진 알림 테스트]\n이 방으로 운영진 알림이 전송됩니다.");
                    noticeStatus.Text = "테스트 알림을 전송 대기열에 추가했습니다.";
                }
                catch (Exception error)
                {
                    string reason = error.GetBaseException().Message;
                    noticeStatus.Text = reason;
                    // 주기적인 상태 표시가 오류를 덮어도 사용자가 실패 이유를 확인할 수 있습니다.
                    MessageBox.Show(this, reason, "테스트 알림 전송 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            retry.Click += (s, e) => { try { if (operatorNotices != null) operatorNotices.RetryBlocked(); } catch (Exception error) { noticeStatus.Text = error.Message; } };
            bot.OperatorHistoryRequested = (command, userId) =>
            {
                if (!config.OperatorAlertsEnabled || !config.AlertHistory || bot.SelectedRoom == null) return;
                string body = Database.Instance.OperatorHistory(command.ChatId, userId, false);
                if (!config.AlertMemoBody)
                { int start = body.IndexOf("[운영 메모]", StringComparison.Ordinal); if (start >= 0) body = body.Substring(0, start) + "운영 메모 본문은 봇 화면에서 확인하세요."; }
                try { QueueNotice("history:" + command.ChatId + ":" + command.LogId, "[" + bot.SelectedRoom.Name + " · 운영진 이력 조회]\n요청자: " + Database.Instance.ChatUserName(command.ChatId, command.AuthorId, command.Nickname) + "\n" + body); }
                catch (Exception error) { noticeStatus.Text = "이력 알림 저장 실패: " + error.GetBaseException().Message; }
            };
            ApplyNoticeRouting();
        }
        private void ApplyNoticeRouting()
        {
            bot.OperatorNotificationAccount = Settings.Instance.OperatorAlertAccount;
            bot.OperatorNotificationRoomId = Settings.Instance.OperatorAlertsEnabled ? Settings.Instance.OperatorAlertRoomId : 0;
        }
        private void LoadNoticeRooms()
        {
            noticeRoom.Items.Clear();
            if (catalog == null) { noticeStatus.Text = "채팅 탭에서 방 목록을 먼저 불러오세요."; return; }
            foreach (var room in catalog.Rooms.Where(r => !r.IsDirect).OrderBy(r => r.Name))
            {
                var item = new NoticeRoomChoice { Room = room }; noticeRoom.Items.Add(item);
                if (room.ChatId == Settings.Instance.OperatorAlertRoomId && room.AccountPath == Settings.Instance.OperatorAlertAccount) noticeRoom.SelectedItem = item;
            }
        }
        private bool SaveNoticeSettings()
        {
            string previousSettings = Newtonsoft.Json.JsonConvert.SerializeObject(Settings.Instance);
            try
            {
                if (operatorNotices == null) throw new InvalidOperationException("알림 기록 오류를 먼저 해결하세요.");
                var choice = noticeRoom.SelectedItem as NoticeRoomChoice;
                if (noticeEnabled.Checked)
                {
                    if (choice == null || bot.SelectedRoom == null) throw new InvalidOperationException("감시 방과 알림방을 각각 선택하세요.");
                    if (choice.Room.ChatId == bot.SelectedRoom.ChatId || choice.Room.AccountPath != bot.SelectedRoom.AccountPath)
                        throw new InvalidOperationException("감시 방과 같은 계정의 다른 방을 선택하세요.");
                    if (catalog.Rooms.Count(r => r.AccountPath == choice.Room.AccountPath && r.Name == choice.Room.Name) != 1)
                        throw new InvalidOperationException("동명 방이 있어 전송 대상을 구분할 수 없습니다. 방 이름을 구분해 주세요.");
                }
                var c = Settings.Instance;
                foreach (var item in operatorNotices.Items.Where(r => r.State == "pending" || r.State == "blocked"))
                { item.State = "cancelled"; item.Error = "알림 설정 저장으로 이전 대기를 취소했습니다."; }
                operatorNotices.Save();
                // 항목을 새로 켰다고 과거 메모까지 소급 전송하지 않습니다.
                c.OperatorAlertSince = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                c.OperatorAlertsEnabled = noticeEnabled.Checked;
                if (choice != null) { c.OperatorAlertAccount = choice.Room.AccountPath; c.OperatorAlertRoomId = choice.Room.ChatId; }
                c.AlertReentry = noticeReentry.Checked; c.AlertMemo = noticeMemo.Checked; c.AlertHistory = noticeHistory.Checked;
                c.AlertMonthly = noticeMonthly.Checked; c.AlertReset = noticeReset.Checked; c.AlertMemoBody = noticeMemoBody.Checked;
                Settings.Save(c); ApplyNoticeRouting(); noticeStatus.Text = "설정 저장 완료. 기존 대기는 취소되며 새 알림부터 적용합니다."; return true;
            }
            catch (Exception error)
            {
                Newtonsoft.Json.JsonConvert.PopulateObject(previousSettings, Settings.Instance); ApplyNoticeRouting();
                noticeStatus.Text = "설정 저장 실패: " + error.GetBaseException().Message; return false;
            }
        }
        private void QueueNotice(string key, string body)
        {
            var c = Settings.Instance;
            if (operatorNotices == null || !c.OperatorAlertsEnabled || bot.SelectedRoom == null || bot.SelectedRoom.AccountPath != c.OperatorAlertAccount || bot.SelectedRoom.ChatId == c.OperatorAlertRoomId)
                throw new InvalidOperationException("알림 기록·사용 설정·감시 방과 계정을 먼저 확인하세요.");
            operatorNotices.Enqueue(key, c.OperatorAlertAccount, c.OperatorAlertRoomId, bot.SelectedRoom.ChatId, body, DateTimeOffset.UtcNow.ToUnixTimeSeconds());
        }
        private void TickOperatorNotifications()
        {
            if (operatorNotices == null) return;
            if (!Database.Instance.UserStorageReady)
            { noticeStatus.Text = "알림 전송 대기: 사용자 DB 초기화가 완료되지 않았습니다. 채팅 탭의 DB 오류를 확인하세요."; return; }
            try
            {
                var c = Settings.Instance;
                if (c.OperatorAlertsEnabled)
                {
                    if (operatorNotices.Status == "운영진 알림 꺼짐") operatorNotices.Status = "운영진 알림 사용 중 · 새 알림 대기";
                    // 등록·재입장 기록 저장이 성공한 뒤에만 외부 전송 대기열을 만듭니다.
                    if (bot.SelectedRoom != null && Database.Instance.HasPendingActivity && Database.Instance.Operations.Rows.Values.Any(r =>
                        r.ChatId == bot.SelectedRoom.ChatId && r.At >= c.OperatorAlertSince &&
                        (r.Kind == "memo" && c.AlertMemo || r.Kind == "reentry" && c.AlertReentry) && !operatorNotices.Items.Any(n => n.Key == r.Key)))
                        Database.Instance.UpdateUserTable();
                    operatorNotices.Collect(Database.Instance, bot.SelectedRoom, c);
                    if (operatorNotices.Items.Any(r => r.State == "pending") && Database.Instance.HasPendingActivity) Database.Instance.UpdateUserTable();
                    ChatRoomInfo destination = null;
                    operatorNotices.Process(c.OperatorAlertAccount, c.OperatorAlertRoomId, item =>
                    {
                        if (bot.SelectedRoom == null || bot.SelectedRoom.ChatId != item.SourceId || bot.SelectedRoom.AccountPath != item.Account) return "감시 방이 변경되었습니다.";
                        if (catalog == null) return "방 목록을 먼저 불러오세요.";
                        var rooms = catalog.Rooms.Where(r => r.ChatId == item.RoomId && r.AccountPath == item.Account).ToArray();
                        if (rooms.Length != 1) return "알림방 ID를 확인할 수 없습니다.";
                        destination = rooms[0];
                        if (destination.ChatId == bot.SelectedRoom.ChatId || catalog.Rooms.Count(r => r.AccountPath == destination.AccountPath && r.Name == destination.Name) != 1) return "알림방이 감시 방이거나 동명 방이 존재합니다.";
                        try { if (!new NativeMentionInput(destination.Name, destination.ProcessId).Read().IsEmpty) return "알림방에 작성 중인 내용이 있습니다."; }
                        catch (Exception error) { return error.GetBaseException().Message; }
                        return null;
                    }, item => WindowsMacro.Instance.SendOperatorNotice(destination.Name, destination.ProcessId, item.Body));
                    noticeStatus.Text = operatorNotices.Status;
                }
                else noticeStatus.Text = "운영진 알림 꺼짐";
                var recent = operatorNotices.Items.OrderByDescending(r => r.At).Take(100).ToArray();
                string signature = string.Join("|", recent.Select(r => r.Key + ":" + r.State + ":" + r.Error));
                if (signature == noticeLogSignature) return;
                noticeLogSignature = signature;
                var selected = noticeLog.SelectedItem as NoticeLogChoice;
                string selectedKey = selected == null ? null : selected.Notice.Key;
                noticeLog.BeginUpdate(); noticeLog.Items.Clear();
                foreach (var item in recent)
                { var choice = new NoticeLogChoice { Notice = item }; noticeLog.Items.Add(choice); if (item.Key == selectedKey) noticeLog.SelectedItem = choice; }
                noticeLog.EndUpdate();
            }
            catch (Exception error) { noticeStatus.Text = "운영진 알림 중단: " + error.GetBaseException().Message; }
        }
    }
}
