using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace KakaotalkBot
{
    internal interface IMentionInput
    {
        InputMentionSnapshot Read();
        void BeginMention(InputMentionSnapshot empty, string search);
        void ChooseCandidate(InputMentionSnapshot seed);
        InputMentionSnapshot Append(InputMentionSnapshot owned, string text);
        InputMentionSnapshot Retarget(InputMentionSnapshot owned, long userId);
        void SendOnce(InputMentionSnapshot ready);
        void Clear(InputMentionSnapshot owned);
    }

    /// <summary>멘션 하나와 본문을 준비하고, 대상 ID를 확인한 후 한 번 전송한다.</summary>
    public sealed class KakaoMentionSender
    {
        private readonly IMentionInput input;
        private readonly Action<int> delay;
        private InputMentionSnapshot prepared;

        public KakaoMentionSender(string roomTitle, int expectedProcessId = 0)
            : this(new NativeMentionInput(roomTitle, expectedProcessId), Thread.Sleep) { }

        internal KakaoMentionSender(IMentionInput input, Action<int> delay) { this.input = input; this.delay = delay; }

        /// <summary>닉네임으로 개인 멘션 후보를 만들고 userId로 대상을 확정한다. 전송하지 않는다.</summary>
        public InputMentionSnapshot Prepare(long userId, string candidateNickname, string body, IEnumerable<string> fallbackNicknames = null)
        {
            if (prepared != null) throw new InvalidOperationException("이미 준비한 메시지를 전송하거나 정리한 뒤 다시 준비하세요.");
            if (userId <= 0) throw new ArgumentOutOfRangeException("userId");
            if (string.IsNullOrWhiteSpace(candidateNickname) || candidateNickname.Length > 128 || candidateNickname.Any(c => char.IsControl(c) || c == '\uFFFC'))
                throw new ArgumentException("멘션 후보를 검색할 닉네임이 필요합니다.", "candidateNickname");
            string search = "@" + candidateNickname;
            if (string.IsNullOrWhiteSpace(body)) throw new ArgumentException("보낼 본문이 필요합니다.", "body");
            body = body.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\n', '\r');
            if (body.Length > 3500 || body.Any(c => c == '\uFFFC' || (char.IsControl(c) && c != '\r' && c != '\t')))
                throw new ArgumentException("본문은 3,500자 이내이며 개행·탭 이외의 제어 문자를 포함할 수 없습니다.", "body");
            InputMentionSnapshot owned = null;
            try
            {
                var empty = input.Read();
                if (!empty.IsEmpty) throw new InvalidOperationException("채팅 입력창에 작성 중인 내용이 있습니다. 비운 뒤 다시 시작하세요.");
                InputMentionSnapshot selected = null;
                // 오픈채팅봇 등이 첫 후보를 가리면 다른 재실 사용자의 검색어로 객체를 만듭니다.
                var candidates = new[] { candidateNickname }.Concat(fallbackNicknames ?? Enumerable.Empty<string>())
                    .Where(name => !string.IsNullOrWhiteSpace(name) && name.Length <= 128 &&
                        !name.Any(c => char.IsControl(c) || c == '\uFFFC'))
                    .Distinct(StringComparer.Ordinal).Take(6).ToArray();
                foreach (string name in candidates)
                {
                    search = "@" + name;
                    input.BeginMention(empty, search);
                    owned = input.Read();
                    if (owned.Text != search + "\r" || owned.Mentions.Count != 0)
                    { owned = null; throw new InvalidOperationException("멘션 후보 입력을 확인하지 못했습니다."); }
                    input.ChooseCandidate(owned);
                    for (int attempt = 0; attempt < 20; attempt++)
                    {
                        var current = input.Read();
                        if (current.Mentions.Count == 1 && current.Mentions[0].Position == 0 &&
                            (current.Text == "\uFFFC \r" || current.Text == "\uFFFC\r")) { selected = current; break; }
                        if (!KakaoInputMentions.SameSnapshot(current, owned))
                        { owned = null; throw new InvalidOperationException("후보 선택 도중 입력 내용이 바뀌었습니다."); }
                        delay(50);
                    }
                    if (selected != null) break;
                    // 직접 입력한 검색어가 그대로 남은 경우에만 지우고 다음 후보를 시도합니다.
                    input.Clear(owned);
                    owned = null;
                    empty = input.Read();
                    if (!empty.IsEmpty) throw new InvalidOperationException("검색어 정리 후 입력 내용이 바뀌었습니다.");
                }
                if (selected == null) throw new InvalidOperationException("일반 사용자 멘션 객체를 만들지 못했습니다. 메시지를 전송하지 않았습니다.");
                owned = selected;
                if (string.Equals(owned.Mentions[0].DisplayText, "@all", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("전체 멘션은 개인 멘션의 원본으로 사용하지 않습니다. 다른 후보 닉네임을 지정하세요.");
                string suffix = (owned.Text == "\uFFFC\r" ? " " : "") + body;
                string expectedText = owned.Text.Substring(0, owned.Text.Length - 1) + suffix + "\r";
                owned = input.Append(owned, suffix);
                if (owned.Text != expectedText || owned.Mentions.Count != 1 || owned.Mentions[0].Position != 0)
                { owned = null; throw new InvalidOperationException("본문을 합친 입력 내용이 예상과 다릅니다."); }
                if (owned.Mentions[0].UserId != userId) owned = input.Retarget(owned, userId);
                if (owned.Mentions.Count != 1 || owned.Mentions[0].UserId != userId || owned.Text != expectedText)
                { owned = null; throw new InvalidOperationException("최종 멘션 대상 검증에 실패했습니다."); }
                prepared = owned;
                return owned;
            }
            catch
            {
                if (owned != null) { try { input.Clear(owned); } catch { /* 사용자 입력이 달라졌으면 지우지 않는다. */ } }
                throw;
            }
        }

        /// <summary>64비트 STA에서 준비와 전송을 한 번 수행한다. 실패 시 자동 재시도하지 않는다.</summary>
        public void Send(long userId, string candidateNickname, string body, IEnumerable<string> fallbackNicknames = null)
        {
            var ready = Prepare(userId, candidateNickname, body, fallbackNicknames);
            SendPrepared(ready);
        }

        /// <summary>이 인스턴스가 준비한 입력을 한 번 전송한다. 완료 판정은 입력창 비워짐까지다.</summary>
        public void SendPrepared(InputMentionSnapshot ready)
        {
            if (ready == null || !ReferenceEquals(ready, prepared)) throw new InvalidOperationException("이 송신기에서 준비한 미전송 메시지가 아닙니다.");
            if (!KakaoInputMentions.SameSnapshot(input.Read(), ready)) throw new InvalidOperationException("준비 이후 입력 내용이 바뀌었습니다. 전송을 취소했습니다.");
            // 전송을 시도한 뒤에는 자동 재시도하거나 사용자 입력을 지우지 않는다.
            prepared = null;
            input.SendOnce(ready);
            for (int attempt = 0; attempt < 20; attempt++)
            {
                delay(50);
                var after = input.Read();
                if (after.IsEmpty) return;
            }
            throw new InvalidOperationException("전송 결과를 확인하지 못했습니다. 채팅방을 확인하기 전에는 다시 보내지 마세요.");
        }

        public void DiscardPrepared(InputMentionSnapshot ready)
        {
            if (ready == null || !ReferenceEquals(ready, prepared)) throw new InvalidOperationException("이 송신기에서 준비한 메시지가 아닙니다.");
            input.Clear(ready);
            prepared = null;
        }
    }

    internal static class KakaoPlainSender
    {
        internal static void Send(IMentionInput input, string body, Action<int> delay)
        {
            if (body == null) throw new ArgumentNullException("body");
            string text = body.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\n', '\r');
            if (string.IsNullOrWhiteSpace(text) || text.Length > 3500 || text.Any(c => c == '\uFFFC' || char.IsControl(c) && c != '\r' && c != '\t'))
                throw new ArgumentException("알림 본문 형식이 올바르지 않습니다.");
            var empty = input.Read();
            if (!empty.IsEmpty) throw new InvalidOperationException("알림방에 작성 중인 내용이 있습니다.");
            // 첫 글자는 기존의 포커스·자리표시자 검증 경로로 입력하고, 나머지는 TOM으로 편집합니다.
            int seedLength = char.IsHighSurrogate(text[0]) ? 2 : 1;
            if (char.IsControl(text[0]) || seedLength > text.Length) throw new ArgumentException("알림 첫 글자가 올바르지 않습니다.");
            input.BeginMention(empty, text.Substring(0, seedLength));
            var ready = input.Read();
            if (ready.Text != text.Substring(0, seedLength) + "\r" || ready.Mentions.Count != 0)
                throw new InvalidOperationException("알림 입력 도중 내용이 바뀌었습니다.");
            if (text.Length > seedLength) ready = input.Append(ready, text.Substring(seedLength));
            if (ready.Text != text + "\r" || ready.Mentions.Count != 0) throw new InvalidOperationException("알림 본문 검증 실패");
            input.SendOnce(ready);
            for (int i = 0; i < 20; i++) { delay(50); if (input.Read().IsEmpty) return; }
            throw new InvalidOperationException("전송 후 입력창이 비워졌는지 확인하지 못했습니다. 자동 재전송하지 않습니다.");
        }

    }

    internal sealed class NativeMentionInput : IMentionInput
    {
        private readonly string title;
        private readonly int expectedPid;
        private InputMentionSnapshot identity;
        public NativeMentionInput(string title, int expectedPid) { this.title = title; this.expectedPid = expectedPid; }

        public InputMentionSnapshot Read()
        {
            var value = KakaoInputMentions.Read(title);
            if (expectedPid != 0 && value.ProcessId != expectedPid) throw new InvalidOperationException("선택한 계정의 카카오톡 창이 아닙니다.");
            if (identity != null && (value.Window != identity.Window || value.ProcessId != identity.ProcessId || value.ProcessStart != identity.ProcessStart))
                throw new InvalidOperationException("준비 중 대상 채팅창이 교체되었습니다.");
            if (identity == null) identity = value;
            return value;
        }

        internal void SendPlain(string body) { KakaoPlainSender.Send(this, body, Thread.Sleep); }

        private void Require(InputMentionSnapshot expected)
        {
            if (!KakaoInputMentions.SameSnapshot(Read(), expected)) throw new InvalidOperationException("입력 내용이 바뀌어 동작을 취소했습니다.");
        }

        public void BeginMention(InputMentionSnapshot empty, string search)
        {
            Require(empty);
            IntPtr room = GetAncestor(empty.Window, 2);
            if (IsIconic(room)) ShowWindow(room, 9);
            SetForegroundWindow(room);
            for (int attempt = 0; attempt < 20 && GetForegroundWindow() != room; attempt++) Thread.Sleep(50);
            if (GetForegroundWindow() != room) throw new InvalidOperationException("채팅방을 활성화할 수 없습니다. 채팅창을 열고 다시 시작하세요.");
            if (!Read().IsEmpty) throw new InvalidOperationException("채팅창을 활성화하는 동안 작성 중인 내용이 생겼습니다.");
            // 빈 입력 컨트롤만 클릭한다. 안내 문구는 첫 문자 입력 때 사라질 수 있다.
            IntPtr point = (IntPtr)((5 << 16) | 5);
            Message(empty.Window, 0x201, (UIntPtr)1, point);
            Message(empty.Window, 0x202, UIntPtr.Zero, point);
            var focused = Read();
            if (!focused.IsEmpty) throw new InvalidOperationException("빈 입력창을 활성화하지 못했습니다.");
            RequireInputFocus(empty.Window);
            // WM_CHAR만 전달하면 카카오톡의 후보 검색 처리가 실행되지 않는다.
            var keys = search.SelectMany(character => new[] { new KeyboardInput { Type = 1, Scan = character, Flags = 4 }, new KeyboardInput { Type = 1, Scan = character, Flags = 6 } }).ToArray();
            if (SendInput((uint)keys.Length, keys, Marshal.SizeOf(typeof(KeyboardInput))) != keys.Length)
                throw new InvalidOperationException("멘션 시작 문자 입력에 실패했습니다.");
            for (int attempt = 0; attempt < 20; attempt++)
            {
                var entered = Read();
                if (entered.Text == search + "\r") return;
                if (!entered.IsEmpty && !((search + "\r").StartsWith(entered.Text.TrimEnd('\r'), StringComparison.Ordinal) && entered.Mentions.Count == 0))
                    throw new InvalidOperationException("멘션 시작 문자 입력 중 내용이 바뀌었습니다.");
                Thread.Sleep(25);
            }
            throw new InvalidOperationException("멘션 시작 문자 입력을 확인하지 못했습니다.");
        }

        public void ChooseCandidate(InputMentionSnapshot seed)
        {
            IntPtr room = GetAncestor(seed.Window, 2);
            IntPtr candidate = IntPtr.Zero;
            for (int attempt = 0; attempt < 20 && candidate == IntPtr.Zero; attempt++)
            {
                Require(seed);
                var matches = new List<IntPtr>();
                EnumWindows((popup, state) =>
                {
                    // 후보 팝업은 자식 창이 아닌 채팅창 소유의 최상위 창이다.
                    if (GetParent(popup) != room || !IsWindowVisible(popup)) return true;
                    var name = new StringBuilder(128); GetClassName(popup, name, name.Capacity);
                    if (name.ToString() != "EVA_Window_Dblclk") return true;
                    EnumChildWindows(popup, (window, parameter) =>
                    {
                        var childName = new StringBuilder(128); GetClassName(window, childName, childName.Capacity);
                        if (GetParent(window) == popup && childName.ToString() == "EVA_ChildWindow_Dblclk" && GetDlgCtrlID(window) == 1148 && IsWindowVisible(window)) matches.Add(window);
                        return true;
                    }, IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
                if (matches.Count > 1) throw new InvalidOperationException("멘션 후보 컨트롤을 유일하게 식별하지 못했습니다.");
                if (matches.Count == 1) candidate = matches[0]; else Thread.Sleep(50);
            }
            if (candidate == IntPtr.Zero) throw new InvalidOperationException("멘션 후보 목록이 열리지 않았습니다.");
            uint pid; GetWindowThreadProcessId(candidate, out pid);
            Rectangle rectangle;
            if (pid != seed.ProcessId || !GetClientRect(candidate, out rectangle) || rectangle.Right < 12 || rectangle.Bottom < 12)
                throw new InvalidOperationException("멘션 후보 컨트롤 상태가 예상과 다릅니다.");
            Require(seed);
            // 후보 선택에 Enter를 쓰지 않는다. 객체 생성은 이후 TOM 조회로 확인한다.
            IntPtr point = (IntPtr)((5 << 16) | 5);
            Message(candidate, 0x201, (UIntPtr)1, point);
            Message(candidate, 0x202, UIntPtr.Zero, point);
            Require(seed);
            Message(candidate, 0x203, (UIntPtr)1, point);
            Message(candidate, 0x202, UIntPtr.Zero, point);
        }

        public InputMentionSnapshot Append(InputMentionSnapshot owned, string text)
        { Require(owned); return KakaoInputMentions.EditOwnedText(owned, text, false); }

        public InputMentionSnapshot Retarget(InputMentionSnapshot owned, long userId)
        {
            Require(owned);
            return KakaoInputMentions.ReplaceUserId(title, 0, owned.Mentions[0].UserId, userId, owned.Text, owned.Mentions[0].DisplayText);
        }

        public void Clear(InputMentionSnapshot owned)
        { Require(owned); KakaoInputMentions.EditOwnedText(owned, "", true); }

        public void SendOnce(InputMentionSnapshot ready)
        {
            Require(ready);
            RequireInputFocus(ready.Window);
            if (!PostMessage(ready.Window, 0x100, (UIntPtr)13, (IntPtr)0x001C0001)) throw new InvalidOperationException("전송 키 전달에 실패했습니다.");
            // 키 누름은 한 번뿐이며, 해제 메시지가 실패해도 다시 누르지 않는다.
            PostMessage(ready.Window, 0x101, (UIntPtr)13, (IntPtr)unchecked((long)0xC01C0001));
        }

        private static void RequireInputFocus(IntPtr window)
        {
            uint pid;
            uint thread = GetWindowThreadProcessId(window, out pid);
            var gui = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            if (GetForegroundWindow() != GetAncestor(window, 2) || !GetGUIThreadInfo(thread, ref gui) || gui.Focus != window)
                throw new InvalidOperationException("입력 포커스가 바뀌었습니다. 전송하지 않았습니다.");
            if (new[] { 0x10, 0x11, 0x12, 0x0D }.Any(key => (GetAsyncKeyState(key) & 0x8000) != 0))
                throw new InvalidOperationException("Shift/Ctrl/Alt/Enter 키가 눌려 있어 입력을 취소했습니다.");
        }

        private static void Message(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam)
        {
            UIntPtr result;
            if (SendMessageTimeout(window, message, wParam, lParam, 3, 1500, out result) == IntPtr.Zero)
                throw new InvalidOperationException("채팅 컨트롤 응답 시간 초과 또는 실패. 전송하지 않았습니다.");
        }
        private delegate bool EnumCallback(IntPtr window, IntPtr state);
        [StructLayout(LayoutKind.Sequential)] private struct Rectangle { public int Left, Top, Right, Bottom; }
        // x64 INPUT의 크기는 마우스 입력 공용체를 포함해 40바이트다.
        [StructLayout(LayoutKind.Explicit, Size = 40)] private struct KeyboardInput
        {
            [FieldOffset(0)] public uint Type;
            [FieldOffset(8)] public ushort Key;
            [FieldOffset(10)] public ushort Scan;
            [FieldOffset(12)] public uint Flags;
        }
        [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint count, KeyboardInput[] input, int size);
        [StructLayout(LayoutKind.Sequential)] private struct GuiThreadInfo
        {
            public int Size, Flags;
            public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            public Rectangle CaretRectangle;
        }
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumCallback callback, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr state);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rectangle rectangle);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool PostMessage(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);
    }
}
