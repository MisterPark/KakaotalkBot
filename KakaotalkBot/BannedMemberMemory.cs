using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KakaotalkBot
{
    internal sealed class BannedMember
    {
        internal long UserId, KickedAt;
        internal string Name;
        internal IntPtr ListWindow;
        internal Rectangle NameBounds;
    }
    // 읽기 전용 프로세스 핸들로 창에 연결된 목록만 읽습니다. 전체 힙 검색·메모리 쓰기는 하지 않습니다.
    internal sealed class BannedMemberMemory : IDisposable
    {
        private readonly IntPtr process;
        internal BannedMemberMemory(int pid)
        {
            process = OpenProcess(0x410, false, pid);
            if (process == IntPtr.Zero) throw new InvalidOperationException("해제 목록 메모리를 읽을 수 없습니다.");
        }
        public void Dispose() { CloseHandle(process); }
        private byte[] Read(ulong address, int length)
        {
            if (address < 0x10000 || address > 0x7FFFFFFFFFFF || length < 1 || length > 65536) return null;
            var bytes = new byte[length]; UIntPtr read;
            return ReadProcessMemory(process, (IntPtr)(long)address, bytes, (UIntPtr)length, out read) && read.ToUInt64() == (ulong)length ? bytes : null;
        }
        private ulong U(ulong address) { var bytes = Read(address, 8); return bytes == null ? 0 : BitConverter.ToUInt64(bytes, 0); }
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
        internal List<BannedMember> ReadList(IntPtr window, long expectedLink)
        {
            ulong layout = (ulong)GetProp(window, "_LayoutManager_Object").ToInt64();
            ulong dialog = 0;
            for (ulong offset = 0; offset < 256; offset += 8)
            {
                ulong complete;
                if (Type(U(layout + offset), out complete) == ".?AVCOpenChatKickedMemberDlg@@") { dialog = complete; break; }
            }
            if (dialog == 0) throw new InvalidOperationException("해제 목록의 소유 객체를 확인하지 못했습니다.");
            ulong control = 0;
            for (ulong offset = 0; offset < 4096; offset += 8)
                if (Is(dialog + offset, ".?AVCKickedMemberListCtrl@@")) { control = dialog + offset; break; }
            if (control == 0) throw new InvalidOperationException("해제 목록 클래스가 변경되었습니다.");
            // 현재 HWND를 실제로 소유한 컨트롤인지 확인합니다.
            bool owns = false;
            for (ulong offset = 0; offset < 256; offset += 8) if (U(control + offset) == (ulong)window.ToInt64()) owns = true;
            if (!owns) throw new InvalidOperationException("해제 목록 창과 객체가 일치하지 않습니다.");
            var candidates = new List<List<BannedMember>>();
            for (ulong offset = 0x100; offset < 0x400; offset += 8)
            {
                ulong head = U(control + offset), count = U(control + offset + 8);
                if (count < 1 || count > 10000) continue;
                var rows = ReadTree(head, (int)count, control, window, expectedLink);
                if (rows != null) candidates.Add(rows);
            }
            if (candidates.Count != 1) throw new InvalidOperationException("해제 목록이 비었거나 목록 구조를 유일하게 확인하지 못했습니다.");
            return candidates[0].OrderByDescending(x => x.KickedAt).ToList();
        }
        private List<BannedMember> ReadTree(ulong head, int count, ulong control, IntPtr window, long expectedLink)
        {
            var header = Read(head, 32);
            if (header == null || header[25] != 1) return null;
            var pending = new Stack<ulong>(); pending.Push(U(head + 8));
            var seen = new HashSet<ulong>(); var result = new List<BannedMember>();
            while (pending.Count > 0)
            {
                ulong node = pending.Pop();
                if (node == head) continue;
                if (node == 0 || !seen.Add(node) || seen.Count > count) return null;
                var bytes = Read(node, 40);
                if (bytes == null || bytes[25] != 0) return null;
                pending.Push(BitConverter.ToUInt64(bytes, 0)); pending.Push(BitConverter.ToUInt64(bytes, 16));
                ulong item = BitConverter.ToUInt64(bytes, 32);
                if (!Is(item, ".?AVListCtrlItem@CEvaListCtrl@@")) return null;
                ulong data = U(item + 0x38);
                if (!Is(data, ".?AVKickedMemberItemData@@") || U(data + 0x10) != item || U(data + 0x18) != control || expectedLink <= 0 || U(data + 0x70) != (ulong)expectedLink) return null;
                ulong user = U(data + 0x78), nameWidget = U(data + 0x60);
                if (!Is(user, ".?AVTalkUser_Mutable@@") || !Is(nameWidget, ".?AVTextWidget@xp@@")) return null;
                ulong id = U(user + 0x20), at = U(data + 0x88), length = U(user + 0x48), capacity = U(user + 0x50);
                if (id == 0 || id > long.MaxValue || at < 1577836800 || at > 4102444800 || length == 0 || length > 1024 || capacity < length || capacity > 4096) return null;
                byte[] name = Read(capacity < 16 ? user + 0x38 : U(user + 0x38), (int)length), rect = Read(nameWidget + 0x80, 16);
                if (name == null || rect == null) return null;
                var bounds = Rectangle.FromLTRB(BitConverter.ToInt32(rect, 0), BitConverter.ToInt32(rect, 4), BitConverter.ToInt32(rect, 8), BitConverter.ToInt32(rect, 12));
                string text;
                try { text = new UTF8Encoding(false, true).GetString(name); } catch (DecoderFallbackException) { return null; }
                result.Add(new BannedMember { UserId = (long)id, KickedAt = (long)at, Name = text, NameBounds = bounds, ListWindow = window });
            }
            if (result.Count != count || result.Select(r => r.UserId).Distinct().Count() != count) return null;
            return result;
        }
        [DllImport("user32.dll", EntryPoint = "GetPropW", CharSet = CharSet.Unicode)] private static extern IntPtr GetProp(IntPtr window, string name);
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll")] private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] bytes, UIntPtr length, out UIntPtr read);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    }
}
