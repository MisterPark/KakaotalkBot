using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace KakaotalkBot
{
    internal static class ChatRoomSearch
    {
        // 현재 카카오톡의 채팅 목록 뷰(1150)와 그 안의 검색 Edit(100)를 찾습니다.
        // 형제 EVA_Window의 순서는 광고/탭 구성에 따라 달라지므로 사용하지 않습니다.
        internal static IntPtr FindChatPane(IntPtr main)
        {
            if (main == IntPtr.Zero || !IsWindow(main)) throw new InvalidOperationException("카카오톡 메인 창을 확인하지 못했습니다.");
            var matches = new List<IntPtr>();
            EnumChildWindows(main, (window, state) =>
            {
                if (GetDlgCtrlID(window) == 1150 && GetAncestor(window, 2) == main) matches.Add(window);
                return true;
            }, IntPtr.Zero);
            if (matches.Count > 1) throw new InvalidOperationException("채팅 목록 영역을 유일하게 확인하지 못했습니다.");
            return matches.Count == 1 ? matches[0] : IntPtr.Zero;
        }

        internal static IntPtr FindSearchEdit(IntPtr main, IntPtr pane)
        {
            if (pane == IntPtr.Zero || !IsWindow(pane) || GetAncestor(pane, 2) != main || GetDlgCtrlID(pane) != 1150)
                throw new InvalidOperationException("채팅 목록 영역을 확인하지 못했습니다.");
            var matches = new List<IntPtr>();
            EnumChildWindows(pane, (window, state) =>
            {
                var name = new StringBuilder(64); GetClassName(window, name, name.Capacity);
                if (GetDlgCtrlID(window) == 100 && name.ToString() == "Edit") matches.Add(window);
                return true;
            }, IntPtr.Zero);
            if (matches.Count > 1) throw new InvalidOperationException("채팅 검색 입력창을 유일하게 확인하지 못했습니다.");
            return matches.Count == 1 ? matches[0] : IntPtr.Zero;
        }

        internal static IntPtr OpenSearch(IntPtr main)
        {
            IntPtr pane = FindChatPane(main);
            for (int attempt = 0; (pane == IntPtr.Zero || !IsWindowVisible(pane)) && attempt < 5; attempt++)
            {
                // 이미 채팅 탭이면 전환하지 않습니다. 다른 탭일 때만 메인 탭을 순환합니다.
                Focus(main, main, false);
                Shortcut(main, 0x09);
                var changed = Stopwatch.StartNew();
                do { Thread.Sleep(50); pane = FindChatPane(main); }
                while ((pane == IntPtr.Zero || !IsWindowVisible(pane)) && changed.ElapsedMilliseconds < 800);
            }
            if (pane == IntPtr.Zero || !IsWindowVisible(pane)) throw new InvalidOperationException("채팅 탭을 열지 못했습니다.");
            IntPtr edit = FindSearchEdit(main, pane);
            if (edit != IntPtr.Zero && IsWindowVisible(edit)) { Focus(main, edit); return edit; }
            Focus(main, pane, false);
            Shortcut(main, 0x46);
            var opened = Stopwatch.StartNew();
            do
            {
                pane = FindChatPane(main);
                edit = FindSearchEdit(main, pane);
                if (edit != IntPtr.Zero && IsWindowVisible(edit)) { Focus(main, edit); return edit; }
                Thread.Sleep(50);
            } while (opened.ElapsedMilliseconds < 1500);
            throw new InvalidOperationException("채팅 검색창을 열지 못했습니다.");
        }

        internal static void ReplaceQuery(IntPtr main, IntPtr edit, string roomName)
        {
            if (string.IsNullOrWhiteSpace(roomName) || roomName.Length > 1024 || roomName.Any(char.IsControl))
                throw new ArgumentException("검색할 채팅방 이름이 올바르지 않습니다.");
            // 전체 교체로 재시도해도 검색어가 쌓이지 않게 합니다. 빈 값부터 적용해 검색 갱신도 유도합니다.
            Write(main, edit, "");
            RequireQuery(main, edit, "");
            Write(main, edit, roomName);
            RequireQuery(main, edit, roomName);
        }

        internal static void RequireQuery(IntPtr main, IntPtr edit, string expected)
        {
            RequireControl(main, edit);
            UIntPtr length;
            if (SendMessageTimeout(edit, 0xE, UIntPtr.Zero, IntPtr.Zero, 3, 1000, out length) == IntPtr.Zero)
                throw new InvalidOperationException("채팅방 검색어 길이를 확인하지 못했습니다.");
            if (length.ToUInt64() != (ulong)expected.Length)
                throw new InvalidOperationException("채팅방 검색어가 일치하지 않아 열기를 중단했습니다.");
            var text = new StringBuilder(expected.Length + 2);
            UIntPtr read;
            if (ReadMessageTimeout(edit, 0xD, (UIntPtr)text.Capacity, text, 3, 1000, out read) == IntPtr.Zero || text.ToString() != expected)
                throw new InvalidOperationException("채팅방 검색어가 일치하지 않아 열기를 중단했습니다.");
        }

        internal static void Submit(IntPtr main, IntPtr edit, string expected)
        {
            RequireQuery(main, edit, expected);
            IntPtr results = WaitForResults(main, edit, expected);
            Focus(main, results);
            RequireQuery(main, edit, expected);
            // Edit에 Enter를 넣으면 처리되지 않은 WM_CHAR로 시스템음이 날 수 있습니다.
            // 검색 결과 컨트롤에만 선택/열기 메시지를 보내며 다른 창에 키를 보내지 않습니다.
            ResultKey(main, results, 0x24); // Home: 이전 검색에서 남은 선택 위치 제거
            RequireQuery(main, edit, expected);
            ResultKey(main, results, 0x0D);
        }

        internal static IntPtr FindResults(IntPtr main, IntPtr pane)
        {
            if (pane == IntPtr.Zero || GetAncestor(pane, 2) != main || GetDlgCtrlID(pane) != 1150)
                throw new InvalidOperationException("채팅 목록 영역을 확인하지 못했습니다.");
            var matches = new List<IntPtr>();
            EnumChildWindows(pane, (window, state) =>
            {
                var name = new StringBuilder(128);
                GetWindowText(window, name, name.Capacity);
                if (GetDlgCtrlID(window) == 1003 && name.ToString().StartsWith("SearchListCtrl_", StringComparison.Ordinal)) matches.Add(window);
                return true;
            }, IntPtr.Zero);
            if (matches.Count > 1) throw new InvalidOperationException("채팅 검색 결과 목록을 유일하게 확인하지 못했습니다.");
            return matches.Count == 1 ? matches[0] : IntPtr.Zero;
        }

        internal static IntPtr WaitForResults(IntPtr main, IntPtr edit, string expected)
        {
            var waiting = Stopwatch.StartNew();
            IntPtr previous = IntPtr.Zero;
            long stableSince = 0;
            do
            {
                RequireReadyMain(main);
                RequireQuery(main, edit, expected);
                IntPtr results = FindResults(main, FindChatPane(main));
                if (results == IntPtr.Zero || !IsWindowVisible(results) || !IsWindowEnabled(results))
                { previous = IntPtr.Zero; stableSince = waiting.ElapsedMilliseconds; }
                else if (results != previous) { previous = results; stableSince = waiting.ElapsedMilliseconds; }
                else if (waiting.ElapsedMilliseconds >= 750 && waiting.ElapsedMilliseconds - stableSince >= 250) return results;
                Thread.Sleep(50);
            } while (waiting.ElapsedMilliseconds < 3000);
            throw new InvalidOperationException("채팅방 검색 결과 목록이 준비되지 않았습니다.");
        }

        private static void ResultKey(IntPtr main, IntPtr results, int key)
        {
            RequireReadyMain(main);
            RequireKeysReleased();
            if (FindResults(main, FindChatPane(main)) != results || !IsWindowVisible(results) || ReadGui(results).Focus != results)
                throw new InvalidOperationException("검색 결과 목록이 바뀌어 열기를 중단했습니다.");
            int scan = (int)MapVirtualKey((uint)key, 0) << 16;
            if (key == 0x24) scan |= 1 << 24;
            // 카카오톡의 메시지 루프가 Enter를 처리하므로 동기 SendMessage로 건너뛰지 않습니다.
            // HWND에 고정된 큐 메시지여서 열기 직후 포커스가 채팅 입력창으로 옮겨가도 전송 키가 들어가지 않습니다.
            if (!PostMessage(results, 0x100, (UIntPtr)key, (IntPtr)(scan | 1)) ||
                !PostMessage(results, 0x101, (UIntPtr)key, (IntPtr)(scan | unchecked((int)0xC0000001))))
                throw new InvalidOperationException("채팅 검색 결과에 열기 요청을 전달하지 못했습니다.");
        }

        private static void Shortcut(IntPtr main, ushort key)
        {
            RequireReadyMain(main);
            RequireKeysReleased();
            // SendKeys의 전역 상태/처리 완료 추정에 의존하지 않고 누름/해제를 한 묶음으로 보냅니다.
            var keys = new[] { KeyInput(0x11, false), KeyInput(key, false), KeyInput(key, true), KeyInput(0x11, true) };
            uint sent = SendInput((uint)keys.Length, keys, Marshal.SizeOf(typeof(Input)));
            if (sent != keys.Length)
            {
                if (sent > 0)
                {
                    var release = new[] { KeyInput(key, true), KeyInput(0x11, true) };
                    SendInput((uint)release.Length, release, Marshal.SizeOf(typeof(Input)));
                }
                throw new InvalidOperationException("채팅 검색 단축키를 전달하지 못했습니다.");
            }
            Thread.Sleep(80);
        }

        private static void RequireKeysReleased()
        {
            if (new[] { 0x01, 0x02, 0x10, 0x11, 0x12, 0x0D, 0x5B, 0x5C }.Any(key => (GetAsyncKeyState(key) & 0x8000) != 0))
                throw new InvalidOperationException("사용자가 누르고 있는 키가 있어 채팅 검색을 보류했습니다.");
        }

        private static GuiThreadInfo ReadGui(IntPtr window)
        {
            uint pid; uint thread = GetWindowThreadProcessId(window, out pid);
            var gui = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            if (thread == 0 || !GetGUIThreadInfo(thread, ref gui))
                throw new InvalidOperationException("카카오톡 입력 상태를 확인하지 못했습니다.");
            return gui;
        }

        private static void RequireReadyMain(IntPtr main)
        {
            if (!IsWindow(main) || !IsWindowEnabled(main) || GetForegroundWindow() != main || (ReadGui(main).Flags & 0x1E) != 0)
                throw new InvalidOperationException("카카오톡 메인 창이 비활성화되었거나 메뉴/대화상자가 열려 검색을 보류했습니다.");
        }

        private static void Focus(IntPtr main, IntPtr control, bool exact = true)
        {
            RequireReadyMain(main);
            RequireKeysReleased();
            if (!IsWindowVisible(control) || !IsWindowEnabled(control) || GetAncestor(control, 2) != main)
                throw new InvalidOperationException("검색할 컨트롤이 준비되지 않았습니다.");
            uint pid; uint target = GetWindowThreadProcessId(control, out pid), current = GetCurrentThreadId();
            bool attached = target != current;
            if (attached && !AttachThreadInput(current, target, true))
                throw new InvalidOperationException("채팅 검색 포커스를 연결하지 못했습니다.");
            try { SetFocus(control); }
            finally { if (attached) AttachThreadInput(current, target, false); }
            RequireReadyMain(main);
            IntPtr focused = ReadGui(control).Focus;
            // 컨테이너는 WM_SETFOCUS를 받은 뒤 내부 목록으로 포커스를 넘길 수 있습니다.
            if (focused != control && (exact || GetAncestor(focused, 2) != main))
                throw new InvalidOperationException("채팅 검색 포커스를 확인하지 못했습니다.");
        }

        private static void Write(IntPtr main, IntPtr edit, string value)
        {
            RequireControl(main, edit);
            UIntPtr result;
            if (WriteMessageTimeout(edit, 0xC, UIntPtr.Zero, value, 3, 1000, out result) == IntPtr.Zero || result == UIntPtr.Zero)
                throw new InvalidOperationException("채팅방 검색어를 교체하지 못했습니다.");
        }

        private static void RequireControl(IntPtr main, IntPtr edit)
        {
            uint mainPid, editPid;
            var name = new StringBuilder(64);
            if (main == IntPtr.Zero || edit == IntPtr.Zero || !IsWindow(main) || !IsWindow(edit) ||
                GetAncestor(edit, 2) != main || GetWindowThreadProcessId(main, out mainPid) == 0 ||
                GetWindowThreadProcessId(edit, out editPid) == 0 || mainPid != editPid ||
                GetClassName(edit, name, name.Capacity) == 0 || name.ToString() != "Edit")
                throw new InvalidOperationException("카카오톡 메인 창에 속한 검색 입력창을 확인하지 못했습니다.");
        }

        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW")] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)] private static extern IntPtr WriteMessageTimeout(IntPtr window, uint message, UIntPtr wParam, string text, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)] private static extern IntPtr ReadMessageTimeout(IntPtr window, uint message, UIntPtr wParam, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
        private delegate bool EnumCallback(IntPtr window, IntPtr parameter);
        [StructLayout(LayoutKind.Sequential)] private struct Rectangle { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct GuiThreadInfo
        {
            public int Size, Flags;
            public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            public Rectangle CaretRectangle;
        }
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumCallback callback, IntPtr state);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint current, uint target, bool attach);
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll")] private static extern uint MapVirtualKey(uint key, uint type);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, UIntPtr key, IntPtr data);
        [StructLayout(LayoutKind.Sequential)] private struct Input { public uint Type; public InputData Data; }
        [StructLayout(LayoutKind.Explicit)] private struct InputData
        {
            [FieldOffset(0)] public KeyboardInput Keyboard;
            [FieldOffset(0)] public MouseInput Mouse;
        }
        [StructLayout(LayoutKind.Sequential)] private struct KeyboardInput { public ushort Key, Scan; public uint Flags, Time; public UIntPtr Extra; }
        [StructLayout(LayoutKind.Sequential)] private struct MouseInput { public int X, Y; public uint Data, Flags, Time; public UIntPtr Extra; }
        private static Input KeyInput(ushort key, bool up)
        { return new Input { Type = 1, Data = new InputData { Keyboard = new KeyboardInput { Key = key, Flags = up ? 2U : 0U } } }; }
        [DllImport("user32.dll")] private static extern uint SendInput(uint count, Input[] inputs, int size);
    }
}
