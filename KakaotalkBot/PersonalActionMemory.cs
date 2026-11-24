using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;

namespace KakaotalkBot
{
    // 현재 창에서 출발하는 읽기 전용 객체 탐색. OCR, 전체 힙 검색, 메모리 쓰기는 사용하지 않습니다.
    internal sealed class PersonalActionMemory : IDisposable
    {
        private readonly IntPtr process;
        private readonly Func<ulong, int, byte[]> reader;
        private readonly Func<IntPtr, string, IntPtr> property;
        internal PersonalActionMemory(int pid) { process = OpenProcess(0x410, false, pid); Require(process != IntPtr.Zero); property = GetProp; }
        internal PersonalActionMemory(Func<ulong, int, byte[]> reader, Func<IntPtr, string, IntPtr> property) { this.reader = reader; this.property = property; }
        public void Dispose() { if (process != IntPtr.Zero) CloseHandle(process); }
        private byte[] Read(ulong at, int length)
        {
            if (at < 65536 || at > 0x7FFFFFFFFFFF || length < 1 || length > 65536) return null;
            if (reader != null) return reader(at, length);
            var bytes = new byte[length]; UIntPtr count;
            return ReadProcessMemory(process, (IntPtr)(long)at, bytes, (UIntPtr)length, out count) && count.ToUInt64() == (ulong)length ? bytes : null;
        }
        private ulong U(ulong at) { var b = Read(at, 8); Require(b != null); return BitConverter.ToUInt64(b, 0); }
        private int I(ulong at) { var b = Read(at, 4); Require(b != null); return BitConverter.ToInt32(b, 0); }
        private string Type(ulong at, out ulong complete)
        {
            complete = 0;
            var p = Read(at, 8); if (p == null) return "";
            ulong table = BitConverter.ToUInt64(p, 0); p = table >= 8 ? Read(table - 8, 8) : null; if (p == null) return "";
            ulong locator = BitConverter.ToUInt64(p, 0); p = Read(locator, 24);
            if (p == null || BitConverter.ToUInt32(p, 0) != 1) return "";
            uint offset = BitConverter.ToUInt32(p, 4), self = BitConverter.ToUInt32(p, 20);
            if (offset > 65536 || at < offset || locator < self) return "";
            var name = Read(locator - self + BitConverter.ToUInt32(p, 12) + 16, 192); if (name == null) return "";
            int end = Array.IndexOf(name, (byte)0); if (end < 0) return "";
            complete = at - offset; return Encoding.ASCII.GetString(name, 0, end);
        }
        private bool Is(ulong at, string name) { ulong ignored; return Type(at, out ignored) == name; }
        internal static void Require(bool valid, [CallerMemberName] string step = "", [CallerLineNumber] int line = 0)
        {
            if (valid) return;
            string label = step == "Pack" ? "이모티콘 팩" : step == "Emoticon" ? "이모티콘 항목" :
                step == "Draft" || step == "DraftController" ? "전송 대기 이모티콘" :
                step == "Reaction" || step == "ReactionContext" ? "공감 대상" :
                step == "EmoticonPicker" ? "이모티콘 선택창" : step == "Room" ? "채팅방" :
                step == "Act" ? "클릭 위치" : "응답 항목";
            throw new InvalidOperationException("카카오톡 " + label + "을 확인하지 못했습니다. 전송하지 않습니다. [" + step + ":" + line + "]");
        }
        private string CString(ulong pointer)
        {
            var bytes = Read(pointer, 128); Require(bytes != null);
            int end = Array.IndexOf(bytes, (byte)0); Require(end >= 0);
            return new UTF8Encoding(false, true).GetString(bytes, 0, end);
        }
        private string StdString(ulong at)
        {
            ulong size = U(at + 16), capacity = U(at + 24); Require(size <= 100 && capacity >= size);
            if (size == 0) return "";
            var bytes = Read(capacity < 16 ? at : U(at), (int)size); Require(bytes != null);
            return new UTF8Encoding(false, true).GetString(bytes);
        }
        internal ulong Dialog(IntPtr window, string name)
        {
            ulong frame = (ulong)property(window, "_CustomWindowFrame_Object").ToInt64();
            ulong root = frame >= 0x570 ? frame - 0x570 : 0;
            return Is(root, ".?AV" + name + "@@") && U(root + 0x40) == (ulong)window.ToInt64() ? root : 0;
        }
        internal ulong Room(IntPtr window, long chat)
        {
            ulong root = Dialog(window, "CChatDlg"); Require(root != 0 && chat > 0);
            ulong vm = U(root + 0x640); Require(Is(vm, ".?AVChatRoom_ViewModel@@") && U(vm + 0xE0) == root);
            ulong reaction = U(vm + 0x370); Require(Is(reaction, ".?AVChatDelegate_Reaction@@") && U(reaction + 0x18) == root && U(reaction + 0x28) == vm);
            ulong shared = U(reaction + 0x60); Require(Is(shared, ".?AV?$_Ref_count_obj2@VChatReactionController@@@std@@") && U(shared + 0x40) == (ulong)chat);
            return root;
        }
        private Rectangle Rect(ulong at)
        {
            var bytes = Read(at, 16); Require(bytes != null);
            var r = Rectangle.FromLTRB(BitConverter.ToInt32(bytes, 0), BitConverter.ToInt32(bytes, 4), BitConverter.ToInt32(bytes, 8), BitConverter.ToInt32(bytes, 12));
            Require(r.Width >= 0 && r.Height >= 0 && r.Width < 4000 && r.Height < 100000); return r;
        }
        private IEnumerable<ulong> Vector(ulong at)
        {
            ulong begin = U(at), end = U(at + 8); Require(end >= begin && end - begin <= 8192 && (end - begin) % 8 == 0);
            var values = new List<ulong>(); for (ulong p = begin; p < end; p += 8) values.Add(U(p));
            Require(U(at) == begin && U(at + 8) == end); return values;
        }
        private List<ulong> Tree(ulong at, int maximum)
        {
            ulong head = U(at), count = U(at + 8); var header = Read(head, 32);
            Require(header != null && header[25] == 1 && count <= (ulong)maximum);
            var nodes = new HashSet<ulong>(); var pending = new Stack<ulong>(); pending.Push(U(head + 8));
            while (pending.Count > 0)
            {
                ulong node = pending.Pop(); if (node == head) continue;
                var bytes = Read(node, 32); Require(bytes != null && bytes[25] == 0 && nodes.Add(node) && (ulong)nodes.Count <= count);
                pending.Push(BitConverter.ToUInt64(bytes, 0)); pending.Push(BitConverter.ToUInt64(bytes, 16));
            }
            Require((ulong)nodes.Count == count && U(at) == head && U(at + 8) == count); return nodes.ToList();
        }
        private List<ulong> Widgets(ulong root)
        {
            var result = new List<ulong>(); var seen = new HashSet<ulong>(); var pending = new Stack<ulong>(); pending.Push(root);
            while (pending.Count > 0)
            {
                ulong current = pending.Pop(), complete; Require(seen.Add(current) && seen.Count < 4096);
                string type = Type(current, out complete); Require(type.EndsWith("@xp@@", StringComparison.Ordinal));
                result.Add(current);
                ulong offset = type == ".?AVTable@xp@@" ? 0x120UL : type == ".?AVTR@xp@@" ? 0x60UL :
                    type == ".?AVTD@xp@@" || type == ".?AVGridWidgetItem@xp@@" ? 0x40UL : type == ".?AVGridWidget@xp@@" ? 0x118UL : 0;
                if (offset == 0) continue;
                foreach (ulong child in Vector(current + offset)) { Require(U(child + 0x28) == current); pending.Push(child); }
            }
            return result;
        }
        private ulong ButtonObject(ulong table, string name)
        {
            var matches = Widgets(table).Where(w => Is(w, ".?AVButtonWidget@xp@@") && CString(U(w + 0x18)) == name).ToArray();
            Require(matches.Length == 1); return matches[0];
        }
        internal Rectangle Button(ulong dialog, string name) { return Rect(ButtonObject(dialog + 0x208, name) + 0x80); }
        private ulong DraftController(ulong room)
        {
            Require(Is(room, ".?AVCChatDlg@@"));
            ulong vm = U(room + 0x640), floating = U(vm + 0x2A8);
            Require(Is(vm, ".?AVChatRoom_ViewModel@@") && U(vm + 0xE0) == room);
            Require(Is(floating, ".?AVChatDelegate_Floating@@") && U(floating + 0x18) == room && U(floating + 0x28) == vm && U(floating + 0x38) == room + 0x650);
            // 입력칸의 인용용 ThumbnailView가 아니라 채팅 로그 하단의 전송 대기 미리보기입니다.
            ulong surface = floating + 0x580;
            Require(Is(surface, ".?AVFloatingDecoView@@") && U(surface + 0x2D0) == room + 0x650);
            ulong controller = U(surface + 0x2F0), shared = U(surface + 0x2F8);
            Require(Is(shared, ".?AV?$_Ref_count_obj2@VChatLogEmoticonDecoController@@@std@@") && controller == shared + 16);
            Require(Is(controller, ".?AVChatLogEmoticonDecoController@@") && U(controller + 0x30) == room && U(controller + 0x38) == U(vm + 0xE8));
            ulong view = U(controller + 0x28);
            Require(Is(view, ".?AVChatLogEmoticonDecoView@@") && U(view + 0x18) == surface + 8 && U(view + 0x28) == controller + 8);
            return controller;
        }
        internal string Draft(ulong room)
        { ulong instance; return Draft(room, out instance); }
        internal string Draft(ulong room, out ulong instance)
        {
            instance = 0;
            ulong controller = DraftController(room), model = U(controller + 0x50), item = U(controller + 0x58);
            // 취소 후 view는 이전 그림을 보관하므로 controller의 활성 참조로 판단합니다.
            if (model == 0 && item == 0) return null;
            Require(Is(model, ".?AVEmoticonObjectModelImpl@@") && Is(item, ".?AUEmoticonExItem@@"));
            Require(U(U(controller + 0x28) + 0x40) == model);
            int index = I(item + 8); string pack = CString(U(item + 0x10));
            Require(index > 0 && index <= 200 && pack.Length > 0 && pack.Length <= 16 && pack.All(c => c >= '0' && c <= '9'));
            string path = CString(U(model + 0x18));
            string identity = ChatMessageDecoder.EmoticonIdentity(20, new Newtonsoft.Json.Linq.JObject { ["path"] = path }.ToString());
            Require(identity == pack + ":" + index && CString(U(model + 0x40)) == pack);
            Require(U(controller + 0x50) == model && U(controller + 0x58) == item);
            instance = model;
            return identity;
        }
        internal Rectangle DraftCancel(ulong room, string expected, out IntPtr window)
        {
            Require(!string.IsNullOrEmpty(expected) && Draft(room) == expected);
            ulong view = U(DraftController(room) + 0x28), log = room + 0x650;
            Require(Is(log, ".?AVCChatLogView@@")); window = (IntPtr)(long)U(log + 0x40);
            return Rect(ButtonObject(U(view + 0x48), "btn_emoticon_preview_cancel") + 0x80);
        }
        private ulong Select(ulong popup)
        {
            Require(Is(popup, ".?AVCEmoticonWnd@@")); ulong select = popup + 0x658; Require(Is(select, ".?AVCEmoticonSelectWnd@@")); return select;
        }
        internal void EmoticonPicker(ulong popup, IntPtr room)
        {
            Require(Is(popup, ".?AVCEmoticonWnd@@") && room != IntPtr.Zero);
            ulong vm = U(popup + 0x630);
            Require(Is(vm, ".?AVEmoticonViewModel@@") && U(vm + 0xE8) == (ulong)room.ToInt64() && U(vm + 0xF0) == (ulong)room.ToInt64());
            Select(popup);
        }
        internal Rectangle Pack(ulong popup, string pack, out IntPtr window)
        {
            ulong title = U(Select(popup) + 0x570); Require(Is(title, ".?AVCEmoticonTitleSelectWnd@@")); window = (IntPtr)(long)U(title + 0x40);
            ulong view = title + 0x5D8; Require(Is(view, ".?AVEmoticonTitleGridView@@")); ulong grid = U(view + 8); Require(Is(grid, ".?AVGridWidget@xp@@"));
            var names = new List<string>();
            foreach (ulong node in Tree(view + 0x78, 1000))
            {
                ulong model = U(node + 0x30), shared = U(node + 0x38);
                Require(Is(shared, ".?AV?$_Ref_count_obj2@VEmoticonExTitle@@@std@@") && model == shared + 16);
                if (CString(U(model + 0x10)) == pack) names.Add(CString(U(node + 0x20)));
            }
            Require(names.Count == 1); return Rect(ButtonObject(grid, names[0]) + 0x80);
        }
        internal Rectangle Emoticon(ulong popup, PersonalEmoticon item, out IntPtr window)
        {
            ulong view = U(Select(popup) + 0x578); Require(Is(view, ".?AVEmoticonEmotsView@@")); window = (IntPtr)(long)U(view + 0x40);
            ulong grid = U(view + 0x5C8), adapter = U(view + 0x5D0), vm = U(view + 0x5E0);
            Require(Is(grid, ".?AVRecycleViewWidget@xp@@") && Is(adapter, ".?AVEmoticonEmotRecycleViewAdapter@@") && U(grid + 0x148) == adapter && U(adapter + 8) == grid + 0x118);
            ulong position = U(grid + 0x140); Require(Is(position, ".?AVListViewPositionManager_Grid@xp@@"));
            var matches = new List<Rectangle>();
            foreach (ulong node in Tree(position + 0x38, 200))
            {
                ulong row = U(node + 0x20), model = U(row + 0x30), shared = U(row + 0x78);
                Require(Is(row, ".?AVRV_ITEM@xp@@") && Is(model, ".?AUEmoticonExItem@@") && model == shared + 16 && U(row + 0x70) == model && Is(shared, ".?AV?$_Ref_count_obj2@UEmoticonExItem@@@std@@"));
                if (I(model + 8) != item.Index || CString(U(model + 0x10)) != item.Pack) continue;
                ulong holder = U(row + 0x60), table = U(row + 0x58), button = U(holder + 0x50);
                Require(Is(holder, ".?AVEmoticonEmotViewHolder@@") && U(holder + 8) == table && U(holder + 0x48) == vm && U(holder + 0x68) == model);
                Require(Is(button, ".?AVButtonWidget@xp@@") && CString(U(button + 0x18)) == "btn_emot_thumbnail");
                var bounds = Rect(button + 0x80); bounds.Offset(-I(grid + 0x130), -I(grid + 0x134)); matches.Add(bounds);
            }
            Require(matches.Count == 1); return matches[0];
        }
        internal void ReactionContext(ulong popup, ulong room, long chat, long log)
        {
            Require(Is(popup, ".?AVReactionPickerWnd@@") && log > 0 && U(popup + 0x638) == (ulong)log);
            ulong reaction = U(U(room + 0x640) + 0x370), shared = U(reaction + 0x60);
            Require(U(popup + 0x5E0) == reaction && U(popup + 0x630) == shared && U(shared + 0x40) == (ulong)chat);
        }
        internal Rectangle Reaction(ulong popup, ulong room, long chat, long log, string code, out IntPtr window)
        {
            ReactionContext(popup, room, chat, log);
            ulong view = popup + 0x650; Require(Is(view, ".?AVReactionPickerContentView@@")); window = (IntPtr)(long)U(view + 0x40);
            ulong grid = U(view + 0x5C0), position = U(grid + 0x140);
            Require(Is(grid, ".?AVRecycleViewWidget@xp@@") && Is(position, ".?AVListViewPositionManager_SectionedGrid@xp@@"));
            var found = new List<Rectangle>();
            foreach (ulong node in Tree(position + 0x38, 1000))
            {
                ulong row = U(node + 0x20), model = U(row + 0x30);
                if (!Is(model, ".?AVReactionPickerItemData@@")) continue;
                if (I(model + 8) != 2 || StdString(model + 0x10) != code) continue;
                ulong holder = U(row + 0x60); Require(Is(holder, ".?AVReactionPickerItemViewHolder@@") && U(holder + 8) == U(row + 0x58) && U(row + 0x70) == model);
                ulong button = U(holder + 0x20); Require(Is(button, ".?AVButtonWidget@xp@@") && CString(U(button + 0x18)) == "btn_reaction_item");
                var bounds = Rect(button + 0x80); bounds.Offset(-I(grid + 0x130), -I(grid + 0x134)); found.Add(bounds);
            }
            Require(found.Count == 1); return found[0];
        }
        private ulong CompleteRow(ulong pointer)
        {
            if (pointer == 0) return 0; ulong complete; Require(Type(pointer, out complete) == ".?AVListCtrlItem@CEvaListCtrl@@"); return complete;
        }
        internal Rectangle Message(ulong room, long chat, long log, long author, bool reactionButton, out IntPtr window)
        {
            ulong control = room + 0x650; Require(Is(control, ".?AVCChatLogView@@")); window = (IntPtr)(long)U(control + 0x40);
            ulong root = U(control + 0x158); var pending = new Stack<Tuple<ulong, ulong, bool>>(); pending.Push(Tuple.Create(root, 0UL, false));
            var seen = new HashSet<ulong>(); ulong target = 0; int top = 0, targetTop = 0, height = 0;
            while (pending.Count > 0)
            {
                var entry = pending.Pop(); ulong row = CompleteRow(entry.Item1); if (row == 0) continue;
                if (entry.Item3)
                {
                    int rowHeight = I(row + 0x18); Require(rowHeight > 0 && rowHeight < 10000);
                    ulong data = U(row + 0x38);
                    if (Is(data, ".?AVChatLogViewMsgLeftItem@@") && U(data + 0x168) == (ulong)chat && U(data + 0x170) == (ulong)log)
                    {
                        ulong user = U(data + 0x3A8); Require(target == 0 && U(data + 0x10) == row && U(data + 0x18) == control && Is(user, ".?AVTalkUser_Mutable@@") && U(user + 0x20) == (ulong)author);
                        target = data; targetTop = top; height = rowHeight;
                    }
                    top += rowHeight; Require(top < 1000000); continue;
                }
                Require(seen.Add(row) && seen.Count <= 5000 && CompleteRow(U(row + 0x20)) == entry.Item2);
                pending.Push(Tuple.Create(U(row + 0x30), row, false)); pending.Push(Tuple.Create(entry.Item1, entry.Item2, true)); pending.Push(Tuple.Create(U(row + 0x28), row, false));
            }
            Require(target != 0 && seen.Count == I(control + 0x160) && top == I(control + 0xF4) && U(control + 0x158) == root);
            int viewport = I(control + 0x100), scrollY = 0;
            if (top > viewport)
            {
                ulong scroll = U(control + 0x338); Require(Is(scroll, ".?AVCCustomScrollCtrl@@") && I(scroll + 0x10C) == top && I(scroll + 0x110) == viewport);
                scrollY = I(scroll + 0x114); Require(scrollY >= 0 && scrollY <= top - viewport);
            }
            var rowBounds = new Rectangle(0, targetTop - scrollY, 3000, height);
            if (!reactionButton) return new Rectangle(72, rowBounds.Top + height / 2 - 3, 6, 6);
            ulong button = ButtonObject(U(target + 0x208), "btn_hovermenu_reaction"); var rect = Rect(button + 0x80);
            Require(rowBounds.Contains(rect) && rect.Width >= 10 && rect.Height >= 10); return rect;
        }
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll")] private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] bytes, UIntPtr length, out UIntPtr count);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetProp(IntPtr window, string name);
    }
}
