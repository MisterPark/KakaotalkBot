using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using KakaotalkBot;
using Newtonsoft.Json.Linq;

internal sealed class LocoMentionProbe : Form
{
    private readonly TextBox email = new TextBox();
    private readonly TextBox password = new TextBox { UseSystemPasswordChar = true };
    private readonly TextBox passcode = new TextBox();
    private readonly TextBox chatId = new TextBox();
    private readonly TextBox userId = new TextBox();
    private readonly TextBox nickname = new TextBox();
    private readonly TextBox body = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical };
    private readonly TextBox output = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly Label status = new Label { AutoSize = true, Text = "시작 중…" };
    private readonly Label account = new Label { AutoSize = true, Text = "로그인 전" };
    private readonly Button login = new Button { Text = "로그인 확인", AutoSize = true };
    private readonly Button requestCode = new Button { Text = "인증번호 요청", AutoSize = true };
    private readonly Button register = new Button { Text = "이 기기 등록", AutoSize = true };
    private readonly Button preview = new Button { Text = "데이터 미리보기", AutoSize = true };
    private readonly Button prepare = new Button { Text = "방·대상 ID 확인", AutoSize = true };
    private readonly Button send = new Button { Text = "확인한 내용 1회 전송", AutoSize = true };
    private readonly Button restart = new Button { Text = "연결 초기화", AutoSize = true };
    private LocoMentionSender sender;
    private string ticket;
    private DateTime expires;
    private bool connected, busy, needsRegistration;
    private bool registrationSupported, authenticationBlocked;
    private readonly Timer timer = new Timer { Interval = 1000 };

    private LocoMentionProbe(bool renderOnly = false)
    {
        Text = "사용자 ID 멘션 · LOCO 호환성 진단";
        Font = new Font("맑은 고딕", 10);
        ClientSize = new Size(880, 820);
        MinimumSize = new Size(750, 720);
        status.MaximumSize = new Size(800, 0);
        StartPosition = FormStartPosition.CenterScreen;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 2, RowCount = 13 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        Controls.Add(layout);
        AddWide(layout, new Label { AutoSize = true, MaximumSize = new Size(800, 0), Text = "카카오톡 입력창을 사용하지 않는 실험용 송신입니다. 현재 서버와의 호환성은 로그인 후 확인합니다.\r\n계정·비밀번호·토큰은 파일에 저장하지 않습니다. 기기 식별자는 exe 아래 LocoSenderData에 저장합니다." }, 0);
        AddRow(layout, "카카오 계정", email, 1);
        AddRow(layout, "비밀번호", password, 2);
        AddWide(layout, Flow(login, restart, account), 3);
        passcode.Width = 120;
        AddRow(layout, "기기 인증번호", Flow(passcode, requestCode, register), 4);
        AddRow(layout, "채팅방 ID", chatId, 5);
        AddRow(layout, "대상 user_id", userId, 6);
        AddRow(layout, "표시 이름", nickname, 7);
        body.Height = 85;
        AddRow(layout, "보낼 본문", body, 8);
        AddWide(layout, Flow(preview, prepare, send), 9);
        AddWide(layout, new Label { AutoSize = true, Text = "전송 전 방 이름·계정 ID·대상 ID·본문을 확인하세요. 연결 오류 후에는 실제 채팅방에서 수신 여부를 확인하세요." }, 10);
        AddWide(layout, output, 11);
        AddWide(layout, status, 12);
        for (int i = 0; i < layout.RowCount; i++) layout.RowStyles.Add(new RowStyle(i == 11 ? SizeType.Percent : SizeType.AutoSize, i == 11 ? 100 : 0));
        foreach (var input in new[] { chatId, userId, nickname, body }) input.TextChanged += (s, e) => { ticket = null; UpdateButtons(); };
        login.Click += async (s, e) => await Run(async () => {
            var result = await sender.LoginAsync(email.Text, password.Text);
            connected = true;
            needsRegistration = false;
            password.Clear();
            passcode.Clear();
            account.Text = "계정 ID: " + result["userId"] + " · 채팅방 " + result["channelCount"] + "개";
            status.Text = "LOCO 로그인 성공. 채팅방과 대상을 확인해 주세요.";
        });
        requestCode.Click += async (s, e) => await Run(async () => {
            await sender.RequestPasscodeAsync(email.Text, password.Text);
            status.Text = "인증번호를 요청했습니다. 카카오톡에서 확인한 번호를 입력해 주세요.";
        });
        register.Click += async (s, e) => await Run(async () => {
            await sender.RegisterDeviceAsync(email.Text, password.Text, passcode.Text);
            passcode.Clear();
            needsRegistration = false;
            status.Text = "기기 등록 완료. 로그인 확인을 눌러 주세요.";
        });
        preview.Click += async (s, e) => await Run(async () => {
            ticket = null;
            output.Text = (await sender.PreviewAsync(ReadId(chatId), Parts())).ToString();
            status.Text = "로컬 데이터 생성만 확인했습니다. 서버 확인과 전송은 하지 않았습니다.";
        });
        prepare.Click += async (s, e) => await Run(async () => {
            ticket = null;
            var result = await sender.PrepareAsync(ReadId(chatId), Parts());
            ticket = (string)result["ticket"];
            expires = DateTime.UtcNow.AddSeconds(110);
            result.Remove("ticket");
            output.Text = result.ToString();
            status.Text = "방·대상 확인 완료. 표시된 내용을 1회 전송할 수 있습니다. (약 2분 유효)";
        });
        send.Click += async (s, e) => await Run(async () => {
            string selectedTicket = ticket;
            ticket = null;
            var result = await sender.SendPreparedAsync(selectedTicket);
            output.AppendText(Environment.NewLine + result.ToString());
            status.Text = "서버가 전송을 수락했습니다. 수신 메시지와 DB의 멘션 user_id를 확인해 주세요.";
        });
        restart.Click += async (s, e) => await Run(StartSender);
        if (!renderOnly) Shown += async (s, e) => await Run(StartSender);
        FormClosed += (s, e) => { timer.Dispose(); password.Clear(); if (sender != null) sender.Dispose(); };
        timer.Tick += (s, e) => UpdateButtons();
        if (!renderOnly) timer.Start();
        UpdateButtons();
    }

    private async Task StartSender()
    {
        if (sender != null) sender.Dispose();
        sender = null;
        connected = false;
        needsRegistration = false;
        registrationSupported = false;
        authenticationBlocked = false;
        ticket = null;
        account.Text = "로그인 전";
        var config = JObject.Parse(File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "runtime.json")));
        sender = new LocoMentionSender((string)config["nodePath"], Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LocoSender", "bridge.js"));
        var info = await sender.GetInfoAsync();
        registrationSupported = (bool?)info["registrationSupported"] == true;
        requestCode.Text = registrationSupported ? "인증번호 요청" : "인증번호 요청 · 미지원";
        register.Text = registrationSupported ? "이 기기 등록" : "기기 등록 · 미지원";
        status.Text = info["library"] + " · 적용 클라이언트 버전 " + info["appVersion"];
        if (!registrationSupported) status.Text += "\r\n" + (string)info["registrationNotice"];
    }

    private LocoMessagePart[] Parts()
    {
        return new[] { LocoMessagePart.Mention(ReadId(userId), nickname.Text), LocoMessagePart.Text(" " + body.Text) };
    }
    private static long ReadId(TextBox control)
    {
        long value;
        if (!long.TryParse(control.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out value) || value <= 0) throw new ArgumentException("ID에는 양의 64비트 정수를 입력해 주세요.");
        return value;
    }

    private async Task Run(Func<Task> operation)
    {
        if (busy) return;
        busy = true;
        status.Text = "처리 중…";
        UpdateButtons();
        try { await operation(); }
        catch (LocoSenderException error)
        {
            ticket = null;
            needsRegistration = error.ServerStatus == -100 || needsRegistration;
            if (error.Code == "REGISTRATION_UNSUPPORTED")
            {
                authenticationBlocked = true;
                needsRegistration = false;
                password.Clear();
                passcode.Clear();
            }
            status.Text = error.Code + (error.ServerStatus.HasValue ? " (" + error.ServerStatus.Value + ")" : "") + ": " + error.Message;
            if (error.Code == "SEND_UNKNOWN" || error.Code == "TIMEOUT" || error.Code == "BRIDGE_CLOSED" || error.Code == "NOT_CONNECTED") connected = false;
        }
        catch (Exception error) { status.Text = error.Message; ticket = null; }
        finally { busy = false; if (!IsDisposed) UpdateButtons(); }
    }

    private void UpdateButtons()
    {
        login.Enabled = !busy && sender != null && !connected && !authenticationBlocked;
        restart.Enabled = !busy;
        requestCode.Enabled = register.Enabled = !busy && sender != null && needsRegistration && registrationSupported;
        passcode.ReadOnly = !registrationSupported;
        preview.Enabled = !busy && sender != null;
        prepare.Enabled = !busy && connected;
        send.Enabled = !busy && connected && ticket != null && expires > DateTime.UtcNow;
        foreach (var input in new[] { email, password, passcode, chatId, userId, nickname, body }) input.Enabled = !busy;
    }

    private static FlowLayoutPanel Flow(params Control[] controls)
    {
        var flow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true };
        flow.Controls.AddRange(controls);
        return flow;
    }
    private static void AddRow(TableLayoutPanel layout, string caption, Control control, int row)
    {
        layout.Controls.Add(new Label { Text = caption, AutoSize = true, Margin = new Padding(3, 9, 3, 8) }, 0, row);
        control.Dock = DockStyle.Fill;
        control.Margin = new Padding(3, 6, 3, 6);
        layout.Controls.Add(control, 1, row);
    }
    private static void AddWide(TableLayoutPanel layout, Control control, int row)
    {
        control.Dock = DockStyle.Fill;
        control.Margin = new Padding(3, 6, 3, 6);
        layout.Controls.Add(control, 0, row);
        layout.SetColumnSpan(control, 2);
    }

    [STAThread]
    private static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        bool renderOnly = args.Length == 1 && args[0] == "--render-check";
        using (var form = new LocoMentionProbe(renderOnly))
        {
            if (renderOnly)
            {
                // 화면 조작 없이 이 폼의 초기 배치를 파일로 렌더링합니다.
                form.ShowInTaskbar = false;
                form.Opacity = 0;
                form.Show();
                form.PerformLayout();
                using (var bitmap = new Bitmap(form.Width, form.Height))
                {
                    form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
                    bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "layout.png"));
                }
                form.timer.Dispose();
                form.Close();
                return;
            }
            Application.Run(form);
        }
    }
}
