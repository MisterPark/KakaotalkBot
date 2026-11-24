using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace KakaotalkBot
{
    internal sealed class ParticipantMember
    {
        internal long UserId;
        internal string Name;
        // Y는 활성 행 트리의 순서/높이로 계산한 내용 좌표입니다. xp 위젯의 Y는 재사용하지 않습니다.
        internal Rectangle AvatarBounds, NameBounds;
        internal int ScrollY, ViewportHeight;
    }
    // 대상 방의 목록 객체만 읽습니다. 프로세스 메모리 쓰기는 하지 않습니다.
    internal sealed class ParticipantMemory : IDisposable
    {
        private readonly IntPtr process;
        private readonly Func<ulong, int, byte[]> readMemory;
        private readonly Func<IntPtr, string, IntPtr> windowProperty;
        internal ParticipantMemory(int pid)
        {
            process = OpenProcess(0x410, false, pid);
            if (process == IntPtr.Zero) throw new InvalidOperationException("참여자 목록 메모리를 읽을 수 없습니다.");
            windowProperty = GetProp;
        }
        internal ParticipantMemory(Func<ulong, int, byte[]> readMemory, Func<IntPtr, string, IntPtr> windowProperty)
        { this.readMemory = readMemory; this.windowProperty = windowProperty; }
        public void Dispose() { if (process != IntPtr.Zero) CloseHandle(process); }
        private byte[] Read(ulong address, int length)
        {
            if (address < 0x10000 || address > 0x7FFFFFFFFFFF || length < 1 || length > 65536) return null;
            if (readMemory != null) return readMemory(address, length);
            var bytes = new byte[length]; UIntPtr read;
            return ReadProcessMemory(process, (IntPtr)(long)address, bytes, (UIntPtr)length, out read) && read.ToUInt64() == (ulong)length ? bytes : null;
        }
        private ulong U(ulong address) { var bytes = Read(address, 8); return bytes == null ? 0 : BitConverter.ToUInt64(bytes, 0); }
        private int I(ulong address) { var bytes = Read(address, 4); return bytes == null ? -1 : BitConverter.ToInt32(bytes, 0); }
        private string Type(ulong address, out ulong complete)
        {
            complete = 0;
            ulong table = U(address), locator = table >= 8 ? U(table - 8) : 0;
            var info = Read(locator, 24);
            if (info == null || BitConverter.ToUInt32(info, 0) != 1) return "";
            uint offset = BitConverter.ToUInt32(info, 4), self = BitConverter.ToUInt32(info, 20);
            if (offset > 65536 || address < offset || locator < self) return "";
            var name = Read(locator - self + BitConverter.ToUInt32(info, 12) + 16, 160);
            if (name == null) return "";
            complete = address - offset;
            int end = Array.IndexOf(name, (byte)0);
            return Encoding.ASCII.GetString(name, 0, end >= 0 ? end : name.Length);
        }
        private bool Is(ulong address, string name) { ulong ignored; return Type(address, out ignored) == name; }
        private ulong Control(IntPtr window)
        {
            ulong layout = (ulong)windowProperty(window, "_LayoutManager_Object").ToInt64(), dialog = 0;
            for (ulong offset = 0; layout != 0 && offset < 256; offset += 8)
            {
                ulong complete;
                if (Type(U(layout + offset), out complete) == ".?AVCGroupInfoWnd@@") { dialog = complete; break; }
            }
            for (ulong offset = 0; dialog != 0 && offset < 4096; offset += 8)
                if (Is(dialog + offset, ".?AVCChatMemberListCtrl@@"))
                {
                    ulong control = dialog + offset;
                    for (ulong field = 0; field < 256; field += 8)
                        if (U(control + field) == (ulong)window.ToInt64()) return control;
                }
            return 0;
        }
        internal bool IsParticipantList(IntPtr window) { return Control(window) != 0; }
        internal void RequireProfile(IntPtr window, IntPtr listWindow, long expectedChat, long expectedUser)
        {
            if (expectedChat <= 0 || expectedUser <= 0) throw new InvalidOperationException("프로필 검증 대상 ID가 유효하지 않습니다.");
            ulong frame = (ulong)windowProperty(window, "_CustomWindowFrame_Object").ToInt64();
            // 관측한 ProfileView_OpenChat의 프레임 멤버에서 실제 창 객체를 복원합니다.
            // RTTI·HWND·목록·모델·컨트롤러의 연결이 모두 맞아야 이 버전의 구조로 인정합니다.
            ulong view = frame >= 0x570 ? frame - 0x570 : 0, control = Control(listWindow);
            if (!Is(view, ".?AVProfileView_OpenChat@@") || U(view + 0x40) != (ulong)window.ToInt64() ||
                control == 0 || U(view + 0x5E0) != control || U(control + 0x3F8) != (ulong)expectedChat)
                throw new InvalidOperationException("열린 프로필의 창·참여자 목록·방 연결을 확인하지 못했습니다.");
            ulong model = U(view + 0x7B8), controller = U(view + 0x7C8), user = U(model + 0x1C8);
            if (!Is(model, ".?AVProfileViewModel_OpenChat@@") || !Is(controller, ".?AVProfileController_OpenChat@@") ||
                U(controller + 0x28) != model || U(controller + 0x78) != view || !Is(user, ".?AVTalkUser_Mutable@@"))
                throw new InvalidOperationException("열린 프로필의 사용자 객체를 확인하지 못했습니다.");
            if (U(model + 0x1B8) != (ulong)expectedChat || U(user + 0x20) != (ulong)expectedUser)
                throw new InvalidOperationException("열린 프로필이 지정한 참여자와 다릅니다. 내보내기를 중단했습니다.");
            if ((ulong)windowProperty(window, "_CustomWindowFrame_Object").ToInt64() != frame ||
                U(view + 0x40) != (ulong)window.ToInt64() || U(view + 0x5E0) != control ||
                U(view + 0x7B8) != model || U(view + 0x7C8) != controller || U(controller + 0x28) != model ||
                U(controller + 0x78) != view || U(model + 0x1C8) != user ||
                U(model + 0x1B8) != (ulong)expectedChat || U(user + 0x20) != (ulong)expectedUser)
                throw new InvalidOperationException("프로필 정보가 바뀌어 내보내기를 중단했습니다.");
        }
        private Rectangle Bounds(ulong widget)
        {
            var rect = Read(widget + 0x80, 16);
            if (rect == null) throw new InvalidOperationException("참여자 위젯 좌표를 읽지 못했습니다.");
            return Rectangle.FromLTRB(BitConverter.ToInt32(rect, 0), BitConverter.ToInt32(rect, 4), BitConverter.ToInt32(rect, 8), BitConverter.ToInt32(rect, 12));
        }
        private int ScrollPosition(ulong control, out int viewportHeight)
        {
            int extent = I(control + 0xF4); viewportHeight = I(control + 0x100);
            if (extent < 0 || extent > 1000000 || viewportHeight < 1 || viewportHeight > 2500)
                throw new InvalidOperationException("참여자 목록의 표시 영역을 확인하지 못했습니다.");
            if (extent <= viewportHeight) return 0;
            ulong scroll = U(control + 0x338);
            if (!Is(scroll, ".?AVCCustomScrollCtrl@@") || I(scroll + 0x10C) != extent || I(scroll + 0x110) != viewportHeight ||
                I(scroll + 0x118) != 0 || I(scroll + 0x11C) != extent - viewportHeight)
                throw new InvalidOperationException("참여자 목록의 스크롤 객체를 확인하지 못했습니다.");
            int position = I(scroll + 0x114);
            if (position < 0 || position > extent - viewportHeight) throw new InvalidOperationException("참여자 목록의 스크롤 위치가 유효하지 않습니다.");
            return position;
        }
        private ulong ItemObject(ulong pointer)
        {
            if (pointer == 0) return 0;
            ulong complete;
            if (Type(pointer, out complete) != ".?AVListCtrlItem@CEvaListCtrl@@")
                throw new InvalidOperationException("참여자 행 연결 객체를 확인하지 못했습니다.");
            // 활성 트리는 두 번째 기반 클래스 포인터를 저장하므로 RTTI로 완전 객체를 복원합니다.
            return complete;
        }
        private Rectangle RowPosition(ulong control, HashSet<ulong> active, ulong target)
        {
            ulong root = U(control + 0x158);
            var pending = new Stack<Tuple<ulong, ulong, bool>>();
            pending.Push(Tuple.Create(root, 0UL, false));
            var snapshots = new Dictionary<ulong, byte[]>();
            int top = 0; Rectangle result = Rectangle.Empty;
            while (pending.Count > 0)
            {
                var entry = pending.Pop(); ulong item = ItemObject(entry.Item1);
                if (item == 0) continue;
                if (entry.Item3)
                {
                    int height = BitConverter.ToInt32(snapshots[item], 0);
                    if (item == target) result = new Rectangle(0, top, 0, height);
                    top += height;
                    if (top > 1000000) throw new InvalidOperationException("참여자 행 높이가 유효하지 않습니다.");
                    continue;
                }
                var bytes = Read(item + 0x18, 32);
                if (!active.Contains(item) || snapshots.ContainsKey(item) || bytes == null ||
                    ItemObject(BitConverter.ToUInt64(bytes, 8)) != entry.Item2 ||
                    BitConverter.ToInt32(bytes, 0) < 1 || BitConverter.ToInt32(bytes, 0) > 500)
                    throw new InvalidOperationException("참여자 표시 순서와 활성 목록이 일치하지 않습니다.");
                snapshots.Add(item, bytes);
                pending.Push(Tuple.Create(BitConverter.ToUInt64(bytes, 24), item, false));
                pending.Push(Tuple.Create(entry.Item1, entry.Item2, true));
                pending.Push(Tuple.Create(BitConverter.ToUInt64(bytes, 16), item, false));
            }
            if (snapshots.Count != active.Count || top != I(control + 0xF4) || U(control + 0x158) != root)
                throw new InvalidOperationException("참여자 표시 순서 또는 전체 높이가 변경되었습니다.");
            foreach (var row in snapshots)
            {
                var current = Read(row.Key + 0x18, 32);
                if (current == null) throw new InvalidOperationException("참여자 행을 다시 확인하지 못했습니다.");
                for (int i = 0; i < current.Length; i++)
                    if (current[i] != row.Value[i]) throw new InvalidOperationException("참여자 표시 순서가 변경되어 다시 확인해야 합니다.");
            }
            return result;
        }
        internal ParticipantMember ReadTarget(IntPtr window, long expectedChat, long expectedUser)
        {
            if (expectedChat <= 0 || expectedUser <= 0) throw new InvalidOperationException("대상 방 또는 참여자 ID가 유효하지 않습니다.");
            ulong control = Control(window);
            if (control == 0) throw new InvalidOperationException("참여자 목록의 소유 객체를 확인하지 못했습니다.");
            ulong head = U(control + 0x368), count = U(control + 0x370);
            int activeCount = I(control + 0x160);
            int viewportHeight; int scrollY = ScrollPosition(control, out viewportHeight);
            var header = Read(head, 32);
            if (count > 10000 || activeCount < 0 || (ulong)activeCount > count || header == null || header[25] != 1) throw new InvalidOperationException("참여자 목록 구조가 변경되었습니다.");
            var pending = new Stack<ulong>(); pending.Push(U(head + 8));
            var nodes = new HashSet<ulong>(); var users = new HashSet<ulong>(); ParticipantMember target = null;
            int activeRead = 0; ulong targetItem = 0; var activeItems = new HashSet<ulong>();
            while (pending.Count > 0)
            {
                ulong node = pending.Pop();
                if (node == head) continue;
                var bytes = Read(node, 40);
                if (!nodes.Add(node) || (ulong)nodes.Count > count || bytes == null || bytes[25] != 0)
                    throw new InvalidOperationException("참여자 목록을 읽는 중 구조가 변경되었습니다.");
                pending.Push(BitConverter.ToUInt64(bytes, 0)); pending.Push(BitConverter.ToUInt64(bytes, 16));
                ulong item = BitConverter.ToUInt64(bytes, 32), data = U(item + 0x38);
                // 캐시에는 해제 후 다른 종류의 객체로 재사용된 주소도 남습니다.
                // 현재 방·컨트롤·항목에 양방향으로 연결된 참여자만 모은 뒤 활성 행 수와 대조합니다.
                if (!Is(item, ".?AVListCtrlItem@CEvaListCtrl@@") || !Is(data, ".?AVCHAT_MEMBER_ITEM_DATA@@") ||
                    U(data + 0x10) != item || U(data + 0x18) != control || U(data + 0xB0) != (ulong)expectedChat)
                    continue;
                activeRead++;
                if (!activeItems.Add(item)) throw new InvalidOperationException("참여자 행 객체가 중복되었습니다.");
                ulong user = U(data + 0x58); var idBytes = Read(user + 0x20, 8);
                if (!Is(user, ".?AVTalkUser_Mutable@@") || idBytes == null)
                    throw new InvalidOperationException("참여자 사용자 객체를 확인하지 못했습니다.");
                ulong id = BitConverter.ToUInt64(idBytes, 0);
                if (id > long.MaxValue) throw new InvalidOperationException("참여자 ID가 유효하지 않습니다.");
                // 카카오톡의 ID 0 항목은 내보내기 대상이 될 수 없지만 다른 사용자의 조회를 막지 않습니다.
                if (id == 0) continue;
                if (!users.Add(id))
                    throw new InvalidOperationException("참여자 ID를 유일하게 확인하지 못했습니다.");
                if (id != (ulong)expectedUser) continue;
                ulong avatar = U(data + 0x70), label = U(data + 0x80), length = U(user + 0x48), capacity = U(user + 0x50);
                if (!Is(avatar, ".?AVButtonWidget@xp@@") || !Is(label, ".?AVTextWidget@xp@@") || length == 0 || length > 1024 || capacity < length || capacity > 4096)
                    throw new InvalidOperationException("대상 참여자의 프로필 버튼을 확인하지 못했습니다.");
                var name = Read(capacity < 16 ? user + 0x38 : U(user + 0x38), (int)length);
                if (name == null) throw new InvalidOperationException("대상 참여자의 이름을 읽지 못했습니다.");
                string text;
                try { text = new UTF8Encoding(false, true).GetString(name); }
                catch (DecoderFallbackException) { throw new InvalidOperationException("대상 참여자의 이름 데이터가 올바르지 않습니다."); }
                target = new ParticipantMember { UserId = (long)id, Name = text, AvatarBounds = Bounds(avatar), NameBounds = Bounds(label), ScrollY = scrollY, ViewportHeight = viewportHeight };
                targetItem = item;
            }
            if (activeRead != activeCount)
                throw new InvalidOperationException("현재 방에 연결된 참여자 수가 목록의 활성 항목 수와 일치하지 않습니다.");
            var row = RowPosition(control, activeItems, targetItem);
            if (target != null)
            {
                if (row.Height < target.AvatarBounds.Height) throw new InvalidOperationException("대상 프로필 버튼이 행 높이를 벗어났습니다.");
                int delta = row.Top + (row.Height - target.AvatarBounds.Height) / 2 - target.AvatarBounds.Top;
                target.AvatarBounds.Offset(0, delta); target.NameBounds.Offset(0, delta);
            }
            if ((ulong)nodes.Count != count || I(control + 0x160) != activeCount ||
                U(control + 0x368) != head || U(control + 0x370) != count)
                throw new InvalidOperationException("참여자 목록이 변경되어 다시 확인해야 합니다.");
            int currentHeight;
            if (ScrollPosition(control, out currentHeight) != scrollY || currentHeight != viewportHeight)
                throw new InvalidOperationException("참여자 목록이 스크롤되어 좌표를 다시 확인해야 합니다.");
            return target;
        }
        [DllImport("user32.dll", EntryPoint = "GetPropW", CharSet = CharSet.Unicode)] private static extern IntPtr GetProp(IntPtr window, string name);
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll")] private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] bytes, UIntPtr length, out UIntPtr read);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    }
}
