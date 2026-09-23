using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace KakaotalkBot
{
    public partial class Form1
    {
        private TabPage operationsPage;
        private TabControl operationsViews;
        private ListBox operationsUsers, operationsAlerts;
        private TextBox operationsSearch, operationsHistory, operationsMemo, operationsStats;
        private DateTimePicker operationsMonth;
        private Label operationsStatus;
        private DateTime nextOperationsTick;
        private string alertsSignature;
        private long operationsRoomId = -1;

        private sealed class UserChoice
        {
            internal User User;
            internal string Name;
            public override string ToString() { return Name + " · " + User.UserId; }
        }
        private sealed class AlertChoice
        {
            internal OperationRecord Row;
            public override string ToString() { return (Row.Read ? "[확인] " : "[새 알림] ") + OperationsStore.LocalTime(Row.At) + " " + Row.Name + " · " + Row.UserId + " / " + Row.Text; }
        }
        private static TextBox ReadBox()
        { return new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill }; }

        private void InitializeOperationsControls()
        {
            operationsPage = new TabPage("운영"); tabControl1.TabPages.Add(operationsPage);
            var tabs = new TabControl { Dock = DockStyle.Fill }; operationsPage.Controls.Add(tabs);
            operationsViews = tabs;
            var history = new TabPage("유저 이력·메모"); var alerts = new TabPage("재입장 알림"); var stats = new TabPage("통계·월간랭킹");
            tabs.TabPages.AddRange(new[] { history, alerts, stats });
            var split = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(790, 600), SplitterDistance = 220 }; history.Controls.Add(split);
            operationsSearch = new TextBox { Dock = DockStyle.Top }; operationsUsers = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            var reload = new Button { Text = "이름 / ID 검색·새로고침", Dock = DockStyle.Bottom, Height = 32 };
            split.Panel1.Controls.Add(operationsUsers); split.Panel1.Controls.Add(operationsSearch); split.Panel1.Controls.Add(reload);
            reload.Click += (s, e) => RefreshOperationsUsers(); operationsSearch.TextChanged += (s, e) => RefreshOperationsUsers();
            operationsUsers.SelectedIndexChanged += (s, e) => ShowOperationsUser();
            operationsHistory = ReadBox(); operationsMemo = new TextBox { Multiline = true, Dock = DockStyle.Bottom, Height = 90, MaxLength = 1000 };
            var save = new Button { Text = "메모 추가 (작성자: 로컬 운영자)", Dock = DockStyle.Bottom, Height = 36 };
            split.Panel2.Controls.Add(operationsHistory); split.Panel2.Controls.Add(operationsMemo); split.Panel2.Controls.Add(save);
            save.Click += (s, e) =>
            {
                try
                {
                    var user = operationsUsers.SelectedItem as UserChoice;
                    if (bot.SelectedRoom == null || user == null) throw new InvalidOperationException("방과 사용자를 선택하세요.");
                    Database.Instance.AddLocalMemo(bot.SelectedRoom.ChatId, user.User.UserId, operationsMemo.Text);
                    operationsMemo.Clear(); ShowOperationsUser(); operationsStatus.Text = "메모 저장 완료";
                }
                catch (Exception error) { operationsStatus.Text = "메모 저장 실패: " + error.GetBaseException().Message; }
            };
            operationsAlerts = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            var read = new Button { Text = "선택 알림 확인 처리", Dock = DockStyle.Bottom, Height = 35 };
            alerts.Controls.Add(operationsAlerts); alerts.Controls.Add(read);
            read.Click += (s, e) =>
            {
                var item = operationsAlerts.SelectedItem as AlertChoice; if (item == null) return;
                try { item.Row.Read = true; Database.Instance.Operations.Dirty = true; Database.Instance.UpdateUserTable(); alertsSignature = null; }
                catch (Exception error) { operationsStatus.Text = "알림 저장 실패: " + error.GetBaseException().Message; }
            };
            operationsAlerts.DoubleClick += (s, e) =>
            {
                var item = operationsAlerts.SelectedItem as AlertChoice; if (item == null) return;
                LoadOperationsHistory(item.Row.UserId);
            };
            operationsStats = ReadBox(); var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 38 };
            operationsMonth = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM", ShowUpDown = true, Width = 110 };
            var show = new Button { Text = "조회", AutoSize = true }; bar.Controls.Add(operationsMonth); bar.Controls.Add(show);
            stats.Controls.Add(operationsStats); stats.Controls.Add(bar);
            show.Click += (s, e) =>
            {
                if (bot.SelectedRoom == null) { operationsStats.Text = "채팅 탭에서 방을 선택하세요."; return; }
                long chat = bot.SelectedRoom.ChatId; string month = operationsMonth.Value.ToString("yyyy-MM");
                operationsStats.Text = (Database.Instance.RoomStatistics(chat, month) + "\n\n" + Database.Instance.Operations.Ranking(chat, month) +
                    "\n\n" + Database.Instance.Operations.Ranking(chat, month, true) + "\n\n" + Database.Instance.LevelRanking(chat)).Replace("\n", "\r\n");
            };
            operationsStatus = new Label { Text = "채팅 탭에서 선택한 방 기준 · 메모 내용은 봇 화면에서 확인", Dock = DockStyle.Bottom, Height = 36 };
            operationsPage.Controls.Add(operationsStatus);
            InitializeOperatorNotifications();
        }
        private void RefreshOperationsUsers()
        {
            var selected = operationsUsers.SelectedItem as UserChoice; long id = selected == null ? 0 : selected.User.UserId;
            operationsUsers.BeginUpdate(); operationsUsers.Items.Clear(); string query = operationsSearch.Text.Trim();
            long chat = bot.SelectedRoom == null ? 0 : bot.SelectedRoom.ChatId;
            foreach (var item in Database.Instance.UserTable.Where(u => Database.Instance.FindRoomUser(chat, u.UserId) != null)
                .Select(u => new UserChoice { User = u, Name = Database.Instance.ChatUserName(chat, u.UserId) })
                .Where(u => query.Length == 0 || u.Name.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0 || u.User.UserId.ToString().Contains(query)).OrderBy(u => u.Name))
            { operationsUsers.Items.Add(item); if (item.User.UserId == id) operationsUsers.SelectedItem = item; }
            operationsUsers.EndUpdate();
        }
        private void ShowOperationsUser()
        {
            var user = operationsUsers.SelectedItem as UserChoice;
            operationsHistory.Text = bot.SelectedRoom == null || user == null ? "방과 사용자를 선택하세요." :
                Database.Instance.OperatorHistory(bot.SelectedRoom.ChatId, user.User.UserId).Replace("\n", "\r\n");
        }
        private void LoadOperationsHistory(long userId)
        {
            if (closing || IsDisposed || Disposing) return;
            operationsSearch.Text = userId.ToString();
            RefreshOperationsUsers();
            if (operationsUsers.Items.Count > 0) operationsUsers.SelectedIndex = 0;
        }

        private void TickOperations()
        {
            if (DateTime.UtcNow < nextOperationsTick || closing) return;
            nextOperationsTick = DateTime.UtcNow.AddSeconds(1);
            try { Database.Instance.MaintainMonth(); }
            catch (Exception error) { bot.DeferProcessing(error); nextOperationsTick = DateTime.UtcNow.AddSeconds(5); operationsStatus.Text = "DB 수신 유지 · 월 초기화 저장 재시도 대기: " + error.GetBaseException().Message; return; }
            TickOperatorNotifications();
            long chat = bot.SelectedRoom == null ? 0 : bot.SelectedRoom.ChatId;
            if (operationsRoomId != chat || operationsUsers.Items.Count == 0 && operationsSearch.Text.Length == 0)
            { operationsRoomId = chat; RefreshOperationsUsers(); ShowOperationsUser(); }
            var alerts = Database.Instance.Operations.Rows.Values.Where(r => r.Kind == "reentry" && r.ChatId == chat).OrderByDescending(r => r.At).ToArray();
            operationsPage.Text = "운영 (새 알림 " + alerts.Count(r => !r.Read) + ")";
            string signature = chat + ":" + alerts.Length + ":" + alerts.Count(r => !r.Read);
            if (signature != alertsSignature)
            { operationsAlerts.Items.Clear(); foreach (var row in alerts) operationsAlerts.Items.Add(new AlertChoice { Row = row }); alertsSignature = signature; ShowOperationsUser(); }

        }
    }
}
