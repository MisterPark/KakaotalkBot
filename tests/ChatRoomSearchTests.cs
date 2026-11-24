using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using KakaotalkBot;

internal static class ChatRoomSearchTests
{
    private static int passed;
    private static void Check(bool value, string label) { if (!value) throw new Exception(label); passed++; }
    private static void Reject(Action action)
    {
        try { action(); } catch (ArgumentException) { passed++; return; } catch (InvalidOperationException) { passed++; return; }
        throw new Exception("Expected search validation failure");
    }
    [STAThread]
    private static int Main()
    {
        try
        {
            using (var one = new SearchWindow())
            using (var two = new SearchWindow())
            {
                string name = "세 글자보다 훨씬 긴 한글 채팅방 😀 & 이름";
                SetWindowText(one.Edit, name + name + name);
                SetWindowText(one.Draft, "사용자가 작성 중인 채팅 초안");
                int changes = one.Changes;
                for (int i = 0; i < 10; i++)
                {
                    ChatRoomSearch.ReplaceQuery(one.Handle, one.Edit, name);
                    Check(Text(one.Edit) == name, "Repeated reopen accumulated or truncated the room name");
                }
                Check(one.Changes >= changes + 20, "Search changes did not notify the parent");
                Check(Text(one.Draft) == "사용자가 작성 중인 채팅 초안", "Unrelated draft was modified");
                ChatRoomSearch.ReplaceQuery(one.Handle, one.Edit, "짧음");
                Check(Text(one.Edit) == "짧음", "Long-to-short replacement left old text");
                ChatRoomSearch.ReplaceQuery(one.Handle, one.Edit, "   내부 공백   ");
                Check(Text(one.Edit) == "   내부 공백   ", "Room name whitespace was changed");
                Reject(() => ChatRoomSearch.ReplaceQuery(one.Handle, two.Edit, name));
                Check(Text(two.Edit) == "", "Another window's edit was modified");
                Reject(() => ChatRoomSearch.ReplaceQuery(IntPtr.Zero, one.Edit, name));
                Reject(() => ChatRoomSearch.ReplaceQuery(one.Handle, IntPtr.Zero, name));
                Reject(() => ChatRoomSearch.ReplaceQuery(one.Handle, one.Label, name));
                foreach (string invalid in new[] { null, "", "   ", "방\r\n이름", "방\0이름" })
                    Reject(() => ChatRoomSearch.ReplaceQuery(one.Handle, one.Edit, invalid));
                Check(Text(one.Edit) == "   내부 공백   ", "Invalid input changed the search");
                SetWindowText(one.Edit, "다른 검색어");
                Reject(() => ChatRoomSearch.RequireQuery(one.Handle, one.Edit, "새로운 검색"));
                SetWindowText(one.Edit, name + "추가된 글자");
                Reject(() => ChatRoomSearch.Submit(one.Handle, one.Edit, name));
                Application.DoEvents();
                Check(Text(one.Edit) == name + "추가된 글자", "Mismatched search was submitted/modified");
                Reject(() => ChatRoomSearch.Submit(one.Handle, one.Edit, name + "추가된 글자"));
                Check(Text(one.Edit) == name + "추가된 글자", "Inactive window submission changed input");
                using (var refusing = new RefusingEdit(one.Edit))
                {
                    Reject(() => ChatRoomSearch.ReplaceQuery(one.Handle, one.Edit, name));
                    Check(Text(one.Edit) == name + "추가된 글자", "Rejected replacement lost prior contents");
                }
                IntPtr destroyed = two.Edit;
                two.Dispose();
                Reject(() => ChatRoomSearch.ReplaceQuery(one.Handle, destroyed, name));
                Check(ChatRoomSearch.FindChatPane(one.Handle) == IntPtr.Zero, "Missing chat pane must not use an arbitrary sibling");
                IntPtr unrelatedSearch = CreateWindowEx(0, "EDIT", "other search", 0x40000080, 0, 0, 100, 20, one.Handle, (IntPtr)100, IntPtr.Zero, IntPtr.Zero);
                IntPtr pane = CreateWindowEx(0, "STATIC", "chat pane", 0x40000000, 0, 0, 100, 50, one.Handle, (IntPtr)1150, IntPtr.Zero, IntPtr.Zero);
                Check(ChatRoomSearch.FindChatPane(one.Handle) == pane, "Chat pane resolved by stable ID instead of sibling order");
                Check(ChatRoomSearch.FindSearchEdit(one.Handle, pane) == IntPtr.Zero, "Hidden/closed search is detected without selecting another search field");
                IntPtr search = CreateWindowEx(0, "EDIT", "", 0x40000080, 0, 0, 100, 20, pane, (IntPtr)100, IntPtr.Zero, IntPtr.Zero);
                Check(ChatRoomSearch.FindSearchEdit(one.Handle, pane) == search, "Search resolved inside chat pane only");
                ChatRoomSearch.ReplaceQuery(one.Handle, search, name);
                Check(Text(search) == name && Text(unrelatedSearch) == "other search", "Pane-scoped replacement preserves unrelated search");
                Reject(() => ChatRoomSearch.FindSearchEdit(one.Handle, one.Label));
                IntPtr duplicate = CreateWindowEx(0, "EDIT", "", 0x40000080, 0, 0, 100, 20, pane, (IntPtr)100, IntPtr.Zero, IntPtr.Zero);
                Reject(() => ChatRoomSearch.FindSearchEdit(one.Handle, pane));
                DestroyWindow(duplicate);
                DestroyWindow(search);
                Check(ChatRoomSearch.FindSearchEdit(one.Handle, pane) == IntPtr.Zero, "Destroyed search handle is not reused");
                search = CreateWindowEx(0, "EDIT", "", 0x40000080, 0, 0, 100, 20, pane, (IntPtr)100, IntPtr.Zero, IntPtr.Zero);
                Check(ChatRoomSearch.FindSearchEdit(one.Handle, pane) == search, "Recreated search field is rediscovered");
                IntPtr normal = CreateWindowEx(0, "STATIC", "ChatRoomListCtrl_old", 0x40000000, 0, 0, 100, 20, pane, (IntPtr)1003, IntPtr.Zero, IntPtr.Zero);
                Check(ChatRoomSearch.FindResults(one.Handle, pane) == IntPtr.Zero, "Normal list is not a search result");
                IntPtr results = CreateWindowEx(0, "STATIC", "SearchListCtrl_new", 0x40000000, 0, 0, 100, 20, pane, (IntPtr)1003, IntPtr.Zero, IntPtr.Zero);
                Check(ChatRoomSearch.FindResults(one.Handle, pane) == results, "Search result resolved independently of sibling order");
                Reject(() => ChatRoomSearch.FindResults(two.Handle, pane));
                duplicate = CreateWindowEx(0, "STATIC", "SearchListCtrl_duplicate", 0x40000000, 0, 0, 100, 20, pane, (IntPtr)1003, IntPtr.Zero, IntPtr.Zero);
                Reject(() => ChatRoomSearch.FindResults(one.Handle, pane));
                DestroyWindow(duplicate);
                DestroyWindow(results);
                Check(ChatRoomSearch.FindResults(one.Handle, pane) == IntPtr.Zero, "Stale result handle is not reused");
                duplicate = CreateWindowEx(0, "STATIC", "other pane", 0x40000000, 0, 0, 100, 50, one.Handle, (IntPtr)1150, IntPtr.Zero, IntPtr.Zero);
                Reject(() => ChatRoomSearch.FindChatPane(one.Handle));
                DestroyWindow(duplicate);
                Check(!IsWindowVisible(one.Handle), "Test window became visible");
            }
            Console.WriteLine("PASS: " + passed + " native edit checks (hidden test windows only; no KakaoTalk input or chat sends)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static string Text(IntPtr window) { var value = new StringBuilder(8192); GetWindowText(window, value, value.Capacity); return value.ToString(); }
    private sealed class SearchWindow : NativeWindow, IDisposable
    {
        internal IntPtr Edit, Draft, Label;
        internal int Changes;
        internal SearchWindow()
        {
            CreateHandle(new CreateParams { Caption = "Hidden search regression fixture", Width = 200, Height = 100 });
            Edit = CreateWindowEx(0, "EDIT", "", 0x40000080, 0, 0, 100, 20, Handle, (IntPtr)1, IntPtr.Zero, IntPtr.Zero);
            Draft = CreateWindowEx(0, "EDIT", "", 0x40000080, 0, 20, 100, 20, Handle, (IntPtr)2, IntPtr.Zero, IntPtr.Zero);
            Label = CreateWindowEx(0, "STATIC", "label", 0x40000000, 0, 40, 100, 20, Handle, (IntPtr)3, IntPtr.Zero, IntPtr.Zero);
            if (Edit == IntPtr.Zero || Draft == IntPtr.Zero || Label == IntPtr.Zero) throw new Exception("Cannot create native edit fixture");
        }
        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x111 && ((message.WParam.ToInt64() >> 16) & 0xFFFF) == 0x300 && message.LParam == Edit) Changes++;
            base.WndProc(ref message);
        }
        public void Dispose() { if (Handle != IntPtr.Zero) DestroyHandle(); }
    }
    private sealed class RefusingEdit : NativeWindow, IDisposable
    {
        internal RefusingEdit(IntPtr handle) { AssignHandle(handle); }
        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0xC) { message.Result = IntPtr.Zero; return; }
            base.WndProc(ref message);
        }
        public void Dispose() { ReleaseHandle(); }
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(uint exStyle, string name, string text, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool SetWindowText(IntPtr window, string text);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
}
