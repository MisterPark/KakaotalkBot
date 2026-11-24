using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using KakaotalkBot;

internal static class BannedMemberMemoryTests
{
    private const long Chat = 18495328425790001L, Link = 123456789L, User = 9007199254740993L;
    private static int passed;
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); passed++; Console.WriteLine("PASS " + message); }
    private static bool Rejected(Fixture fixture, long chat = Chat)
    { try { using (var reader = fixture.Reader()) reader.ReadList(fixture.Window, chat); return false; } catch (InvalidOperationException) { return true; } }
    private sealed class Memory
    {
        private readonly Dictionary<ulong, byte> bytes = new Dictionary<ulong, byte>();
        private readonly Dictionary<string, ulong> tables = new Dictionary<string, ulong>();
        private ulong next = 0x100000;
        internal ulong Allocate(int size)
        {
            ulong address = next; next += (ulong)((size + 4095) / 4096) * 4096;
            for (int i = 0; i < size; i++) bytes[address + (ulong)i] = 0;
            return address;
        }
        internal void Put(ulong address, byte[] value) { for (int i = 0; i < value.Length; i++) bytes[address + (ulong)i] = value[i]; }
        internal void U(ulong address, ulong value) { Put(address, BitConverter.GetBytes(value)); }
        internal void I(ulong address, int value) { Put(address, BitConverter.GetBytes(value)); }
        internal void Byte(ulong address, byte value) { bytes[address] = value; }
        internal byte[] Read(ulong address, int count)
        {
            var result = new byte[count];
            for (int i = 0; i < count; i++) if (!bytes.TryGetValue(address + (ulong)i, out result[i])) return null;
            return result;
        }
        internal void Type(ulong address, string name)
        {
            ulong table;
            if (!tables.TryGetValue(name, out table))
            {
                ulong module = Allocate(4096), locator = module + 0x100; table = module + 0x300;
                U(table - 8, locator); I(locator, 1); I(locator + 12, 0x200); I(locator + 20, 0x100);
                Put(module + 0x210, Encoding.ASCII.GetBytes(name + "\0")); tables[name] = table;
            }
            U(address, table);
        }
        internal ulong Object(string type, int size) { ulong address = Allocate(size); Type(address, type); return address; }
    }
    private sealed class Fixture
    {
        internal readonly Memory Memory = new Memory();
        internal readonly IntPtr Window = new IntPtr(0x2222);
        internal readonly ulong Layout, Dialog, Control, Head;
        internal readonly List<ulong> Nodes = new List<ulong>(), Items = new List<ulong>(), Data = new List<ulong>(), Users = new List<ulong>();
        internal Fixture(int count = 1, bool longName = false)
        {
            Layout = Memory.Allocate(256); Dialog = Memory.Object(".?AVCOpenChatKickedMemberDlg@@", 4096); Control = Dialog + 0x6E8;
            Memory.U(Layout + 0x40, Dialog); Memory.Type(Control, ".?AVCKickedMemberListCtrl@@"); Memory.U(Control + 0x40, (ulong)Window.ToInt64());
            Head = EmptyHead(); Memory.U(Control + 0x368, Head); Memory.U(Control + 0x370, (ulong)count);
            // 목록과 무관한 빈 트리가 있어도 이를 실제 빈 목록으로 판정하면 안 됩니다.
            Memory.U(Control + 0x120, EmptyHead());
            for (int i = 0; i < count; i++)
            {
                ulong node = Memory.Allocate(40), item = Memory.Object(".?AVListCtrlItem@CEvaListCtrl@@", 128), data = Memory.Object(".?AVKickedMemberItemData@@", 256);
                ulong user = Memory.Object(".?AVTalkUser_Mutable@@", 256), widget = Memory.Object(".?AVTextWidget@xp@@", 160);
                Nodes.Add(node); Items.Add(item); Data.Add(data); Users.Add(user);
                Memory.U(node, Head); Memory.U(node + 8, i == 0 ? Head : Nodes[i - 1]); Memory.U(node + 16, Head); Memory.U(node + 32, item);
                if (i > 0) Memory.U(Nodes[i - 1] + 16, node);
                Memory.U(item + 0x38, data); Memory.U(data + 0x10, item); Memory.U(data + 0x18, Control); Memory.U(data + 0x60, widget);
                Memory.U(data + 0x70, (ulong)Chat); Memory.U(data + 0x78, user); Memory.U(data + 0x88, (ulong)(1700000000L + i));
                Memory.U(user + 0x20, (ulong)(User + i));
                byte[] name = Encoding.UTF8.GetBytes(longName ? "긴이름테스트가나다" : "검증");
                Memory.U(user + 0x48, (ulong)name.Length); Memory.U(user + 0x50, name.Length < 16 ? 15UL : 63UL);
                if (name.Length < 16) Memory.Put(user + 0x38, name);
                else { ulong text = Memory.Allocate(64); Memory.Put(text, name); Memory.U(user + 0x38, text); }
                Memory.I(widget + 0x80, 10); Memory.I(widget + 0x84, 20 + 50 * i); Memory.I(widget + 0x88, 100); Memory.I(widget + 0x8C, 40 + 50 * i);
            }
            if (count > 0) { Memory.U(Head, Nodes[0]); Memory.U(Head + 8, Nodes[0]); Memory.U(Head + 16, Nodes[count - 1]); }
        }
        private ulong EmptyHead()
        {
            ulong head = Memory.Allocate(32); Memory.U(head, head); Memory.U(head + 8, head); Memory.U(head + 16, head); Memory.Byte(head + 25, 1); return head;
        }
        internal BannedMemberMemory Reader()
        { return new BannedMemberMemory(Memory.Read, (window, key) => window == Window && key == "_LayoutManager_Object" ? (IntPtr)(long)Layout : IntPtr.Zero); }
    }
    private static int Main()
    {
        try
        {
            var normal = new Fixture(2);
            using (var reader = normal.Reader())
            {
                var rows = reader.ReadList(normal.Window, Chat);
                Check(rows.Count == 2 && rows[0].UserId == User + 1 && rows[1].UserId == User, "채팅방 ID로 목록 읽기·64비트 사용자 ID 보존·시간 역순");
                Check(rows[1].Name == "검증" && rows[1].NameBounds == new Rectangle(10, 20, 90, 20), "UTF-8 이름과 해제 대상 좌표 읽기");
                Check(reader.IsBanListWindow(normal.Window) && !reader.IsBanListWindow(new IntPtr(999)), "목록 창과 일반 채팅 컨트롤 구분");
            }
            Check(Rejected(normal, Link), "링크 ID를 채팅방 ID로 전달하던 기존 오류 재현");
            Check(Rejected(normal, Chat + 1) && Rejected(normal, 0), "다른 방·미확인 방 ID 거부");
            var heapName = new Fixture(1, true);
            using (var reader = heapName.Reader()) Check(reader.ReadList(heapName.Window, Chat).Single().Name == "긴이름테스트가나다", "긴 UTF-8 이름 읽기");
            var empty = new Fixture(0);
            using (var reader = empty.Reader()) Check(reader.ReadList(empty.Window, Chat).Count == 0, "검증된 항목 트리가 비어 있으면 정상 빈 목록 반환");
            var corruptEmpty = new Fixture(0); corruptEmpty.Memory.U(corruptEmpty.Head + 8, 0);
            Check(Rejected(corruptEmpty), "손상된 빈 트리를 정상 빈 목록으로 오인하지 않음");
            var unknownEmpty = new Fixture(0); unknownEmpty.Memory.U(unknownEmpty.Control + 0x368, 0);
            Check(Rejected(unknownEmpty), "관측되지 않은 빈 목록 구조는 읽기 실패로 구분");
            var otherRoom = new Fixture(1); otherRoom.Memory.U(otherRoom.Data[0] + 0x70, (ulong)(Chat + 1));
            Check(Rejected(otherRoom), "일반 빈 트리가 있어도 방 ID 불일치는 빈 목록으로 처리하지 않음");
            var wrongOwner = new Fixture(); wrongOwner.Memory.U(wrongOwner.Control + 0x40, 0x3333);
            Check(Rejected(wrongOwner), "창과 컨트롤 소유 관계 불일치 거부");
            var wrongControl = new Fixture(); wrongControl.Memory.U(wrongControl.Data[0] + 0x18, wrongControl.Control + 8);
            Check(Rejected(wrongControl), "행의 목록 컨트롤 역참조 검사");
            var wrongItem = new Fixture(); wrongItem.Memory.U(wrongItem.Data[0] + 0x10, wrongItem.Items[0] + 8);
            Check(Rejected(wrongItem), "행 항목 역참조 검사");
            var duplicate = new Fixture(2); duplicate.Memory.U(duplicate.Users[1] + 0x20, (ulong)User);
            Check(Rejected(duplicate), "중복 사용자 ID가 있는 목록 거부");
            var cycle = new Fixture(); cycle.Memory.U(cycle.Nodes[0], cycle.Nodes[0]);
            Check(Rejected(cycle), "트리 순환에서 종료");
            var mismatch = new Fixture(); mismatch.Memory.U(mismatch.Control + 0x370, 2);
            Check(Rejected(mismatch), "목록 개수와 읽은 행 수 불일치 거부");
            var ambiguous = new Fixture(); ambiguous.Memory.U(ambiguous.Control + 0x200, ambiguous.Head); ambiguous.Memory.U(ambiguous.Control + 0x208, 1);
            Check(Rejected(ambiguous), "목록 트리 후보가 둘이면 임의 선택하지 않음");
            var invalidText = new Fixture(); invalidText.Memory.U(invalidText.Users[0] + 0x48, 1); invalidText.Memory.Byte(invalidText.Users[0] + 0x38, 0xC3);
            Check(Rejected(invalidText), "손상된 UTF-8 이름 거부");
            Console.WriteLine("PASS: " + passed + " banned-list checks (synthetic memory, no UI or network)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
