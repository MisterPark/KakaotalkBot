using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using KakaotalkBot;

internal static class ModerationParticipantSearchTests
{
    private const long Chat = 18495328425790001L, User = 9007199254740993L;
    private static int passed;
    private static void Check(bool ok, string message) { if (!ok) throw new Exception(message); passed++; Console.WriteLine("PASS " + message); }
    private sealed class Memory
    {
        internal readonly Dictionary<ulong, byte> Bytes = new Dictionary<ulong, byte>();
        private readonly Dictionary<string, ulong> tables = new Dictionary<string, ulong>();
        private ulong next = 0x100000;
        internal ulong Allocate(int size)
        {
            ulong address = next; next += (ulong)((size + 4095) / 4096) * 4096;
            for (int i = 0; i < size; i++) Bytes[address + (ulong)i] = 0;
            return address;
        }
        internal void Put(ulong address, byte[] value) { for (int i = 0; i < value.Length; i++) Bytes[address + (ulong)i] = value[i]; }
        internal void U(ulong address, ulong value) { Put(address, BitConverter.GetBytes(value)); }
        internal void I(ulong address, int value) { Put(address, BitConverter.GetBytes(value)); }
        internal byte[] Read(ulong address, int count)
        {
            var result = new byte[count];
            for (int i = 0; i < count; i++) if (!Bytes.TryGetValue(address + (ulong)i, out result[i])) return null;
            return result;
        }
        internal void Type(ulong address, string name, int completeOffset = 0)
        {
            ulong table;
            string key = name + ":" + completeOffset;
            if (!tables.TryGetValue(key, out table))
            {
                ulong module = Allocate(4096), locator = module + 0x100; table = module + 0x300;
                U(table - 8, locator); I(locator, 1); I(locator + 4, completeOffset); I(locator + 12, 0x200); I(locator + 20, 0x100);
                Put(module + 0x210, Encoding.ASCII.GetBytes(name + "\0")); tables[key] = table;
            }
            U(address, table);
        }
        internal ulong Object(string name, int size) { ulong address = Allocate(size); Type(address, name); return address; }
        internal void Rect(ulong address, Rectangle value) { I(address, value.Left); I(address + 4, value.Top); I(address + 8, value.Right); I(address + 12, value.Bottom); }
    }
    private sealed class Fixture
    {
        internal readonly Memory M = new Memory();
        internal readonly IntPtr Window = new IntPtr(0x2222);
        internal readonly ulong Layout, Dialog, Control, Head, Scroll;
        internal readonly List<ulong> Nodes = new List<ulong>(), Items = new List<ulong>(), Data = new List<ulong>(), Users = new List<ulong>(), Avatars = new List<ulong>(), Labels = new List<ulong>();
        internal readonly IntPtr ProfileWindow = new IntPtr(0x3333);
        internal ulong Profile, ProfileModel, ProfileController;
        internal Fixture(int count = 2, string name = "ㅇㅇ")
        {
            Layout = M.Allocate(256); Dialog = M.Object(".?AVCGroupInfoWnd@@", 4096); Control = Dialog + 0xAF0;
            M.U(Layout + 0x40, Dialog); M.Type(Control, ".?AVCChatMemberListCtrl@@"); M.U(Control + 0x40, (ulong)Window.ToInt64());
            M.U(Control + 0x3F8, (ulong)Chat);
            Head = M.Allocate(32); M.U(Head, Head); M.U(Head + 8, Head); M.U(Head + 16, Head); M.Bytes[Head + 25] = 1;
            M.U(Control + 0x368, Head); M.U(Control + 0x370, (ulong)count);
            M.I(Control + 0x160, count);
            M.I(Control + 0xF4, count * 52); M.I(Control + 0x100, 280);
            Scroll = M.Object(".?AVCCustomScrollCtrl@@", 512); M.U(Control + 0x338, Scroll);
            M.I(Scroll + 0x10C, count * 52); M.I(Scroll + 0x110, 280); M.I(Scroll + 0x11C, Math.Max(0, count * 52 - 280));
            for (int i = 0; i < count; i++)
            {
                ulong node = M.Allocate(40), item = M.Object(".?AVListCtrlItem@CEvaListCtrl@@", 128), data = M.Object(".?AVCHAT_MEMBER_ITEM_DATA@@", 224);
                ulong user = M.Object(".?AVTalkUser_Mutable@@", 256), avatar = M.Object(".?AVButtonWidget@xp@@", 160), label = M.Object(".?AVTextWidget@xp@@", 160);
                Nodes.Add(node); Items.Add(item); Data.Add(data); Users.Add(user); Avatars.Add(avatar); Labels.Add(label);
                M.Type(item + 8, ".?AVListCtrlItem@CEvaListCtrl@@", 8); M.I(item + 0x18, 52);
                M.U(node, Head); M.U(node + 8, i == 0 ? Head : Nodes[i - 1]); M.U(node + 16, Head); M.U(node + 32, item);
                if (i > 0) M.U(Nodes[i - 1] + 16, node);
                M.U(item + 0x38, data); M.U(data + 0x10, item); M.U(data + 0x18, Control); M.U(data + 0x58, user);
                M.U(data + 0x70, avatar); M.U(data + 0x80, label); M.U(data + 0xB0, (ulong)Chat); M.U(user + 0x20, (ulong)(User + i));
                byte[] text = Encoding.UTF8.GetBytes(name); M.U(user + 0x48, (ulong)text.Length); M.U(user + 0x50, text.Length < 16 ? 15UL : 127UL);
                if (text.Length < 16) M.Put(user + 0x38, text);
                else { ulong storage = M.Allocate(128); M.Put(storage, text); M.U(user + 0x38, storage); }
                M.Rect(avatar + 0x80, new Rectangle(18, 6 + i * 52, 40, 40)); M.Rect(label + 0x80, new Rectangle(70, 18 + i * 52, 180, 17));
            }
            if (count > 0) { M.U(Head, Nodes[0]); M.U(Head + 8, Nodes[0]); M.U(Head + 16, Nodes[count - 1]); }
            ActiveOrder(Items.ToArray());
        }
        internal void ActiveOrder(params ulong[] order)
        {
            M.U(Control + 0x158, order.Length == 0 ? 0 : order[0] + 8);
            for (int i = 0; i < order.Length; i++)
            { M.U(order[i] + 0x20, i == 0 ? 0 : order[i - 1] + 8); M.U(order[i] + 0x28, 0); M.U(order[i] + 0x30, i + 1 == order.Length ? 0 : order[i + 1] + 8); }
            M.I(Control + 0x160, order.Length); M.I(Control + 0xF4, order.Length * 52);
            M.I(Scroll + 0x10C, order.Length * 52); M.I(Scroll + 0x11C, Math.Max(0, order.Length * 52 - 280));
        }
        internal ParticipantMemory Reader(Func<ulong, int, byte[]> read = null)
        { return new ParticipantMemory(read ?? M.Read, (window, key) => window == Window && key == "_LayoutManager_Object" ? (IntPtr)(long)Layout : window == ProfileWindow && key == "_CustomWindowFrame_Object" && Profile != 0 ? (IntPtr)(long)(Profile + 0x570) : IntPtr.Zero); }
        internal void OpenProfile(int member = 0)
        {
            Profile = M.Object(".?AVProfileView_OpenChat@@", 4096); ProfileModel = M.Object(".?AVProfileViewModel_OpenChat@@", 1024);
            ProfileController = M.Object(".?AVProfileController_OpenChat@@", 256);
            M.U(Profile + 0x40, (ulong)ProfileWindow.ToInt64()); M.U(Profile + 0x5E0, Control);
            M.U(Profile + 0x7B8, ProfileModel); M.U(Profile + 0x7C8, ProfileController);
            M.U(ProfileModel + 0x1B8, (ulong)Chat); M.U(ProfileModel + 0x1C8, Users[member]);
            M.U(ProfileController + 0x28, ProfileModel); M.U(ProfileController + 0x78, Profile);
        }
        internal void VerifyProfile(long user = User, long chat = Chat)
        { using (var reader = Reader()) reader.RequireProfile(ProfileWindow, Window, chat, user); }
        internal ParticipantMember Target(long user = User, long chat = Chat)
        { using (var reader = Reader()) return reader.ReadTarget(Window, chat, user); }
    }
    private static bool Rejected(Action action)
    { try { action(); return false; } catch (InvalidOperationException) { return true; } }
    private static int Main()
    {
        try
        {
            var f = new Fixture(); var viewport = new Rectangle(1, 53, 268, 280);
            var target = f.Target();
            Check(target.UserId == User && target.Name == "ㅇㅇ", "native ID and Korean consonant nickname are read without OCR");
            Check(f.Target(User + 1).UserId == User + 1, "same-name rows are resolved by exact 64-bit ID");
            Check(target.AvatarBounds == new Rectangle(18, 6, 40, 40), "profile button geometry is read from the target widget");
            Check(!ModerationParticipantSearch.Target(target, User, "ㅇㅇ", viewport).IsEmpty, "confirmed ID and visible native button can be selected");
            Check(ModerationParticipantSearch.Target(target, User + 1, "ㅇㅇ", viewport).IsEmpty, "different ID is rejected even with the same nickname");
            Check(ModerationParticipantSearch.Target(target, User, "0 0", viewport).IsEmpty, "lookalike nickname is not substituted");
            Check(f.Target(User + 99) == null, "missing ID does not fall back to another matching name");
            Check(new Fixture(0).Target() == null, "loading or empty tree produces no target");
            f = new Fixture(3); f.M.U(f.Users[0] + 0x20, 0);
            Check(f.Target(User + 1).UserId == User + 1, "ID zero entry cannot block a different user's lookup");
            f.M.U(f.Items[2], 0); f.M.U(f.Items[2] + 0x38, 0); f.M.I(f.Control + 0x160, 2);
            f.ActiveOrder(f.Items.Take(2).ToArray());
            Check(f.Target(User + 1).UserId == User + 1, "ID zero row and detached cache entry reproduce the actual failing list");
            f.M.I(f.Control + 0x160, 3);
            Check(Rejected(() => f.Target(User + 1)), "unreadable active row cannot be hidden as a detached cache entry");
            f = new Fixture(105);
            f.M.I(f.Control + 0x160, 84); f.M.I(f.Control + 0xF4, 84 * 52);
            f.M.I(f.Scroll + 0x10C, 84 * 52); f.M.I(f.Scroll + 0x11C, 84 * 52 - 280);
            f.ActiveOrder(f.Items.Take(84).ToArray());
            for (int i = 84; i < 105; i++)
            {
                f.M.U(f.Users[i] + 0x20, (ulong)(User + 15));
                if (i % 4 == 0) { f.M.Type(f.Items[i], ".?AV_EL_ITEM_@@"); f.M.Type(f.Data[i], ".?AVIRefCountable@eva@@"); }
                else if (i % 4 == 1) f.M.Type(f.Data[i], ".?AVChatLogViewItem_FeedMsg@@");
                else if (i % 4 == 2) f.M.U(f.Data[i] + 0xB0, (ulong)(Chat + 1));
                else f.M.U(f.Data[i] + 0x18, f.Control + 8);
            }
            Check(f.Target(User + 15).UserId == User + 15, "live regression: 84 active rows plus 21 reused cache objects still resolve the exact target");
            Check(f.Target(User + 99) == null, "unrelated cached objects cannot supply a missing target");
            f.M.U(f.Nodes[104] + 32, 0x7FFF0000);
            Check(f.Target(User + 15).UserId == User + 15, "freed cache address is excluded only when all active rows are accounted for");
            f.M.U(f.Data[0] + 0xB0, (ulong)(Chat + 1));
            Check(Rejected(() => f.Target(User + 15)), "one invalid active row still blocks selection even when many stale objects are expected");
            f.M.U(f.Data[0] + 0xB0, (ulong)Chat);
            f.M.Type(f.Items[84], ".?AVListCtrlItem@CEvaListCtrl@@"); f.M.Type(f.Data[84], ".?AVCHAT_MEMBER_ITEM_DATA@@");
            f.M.U(f.Users[84] + 0x20, (ulong)(User + 84));
            Check(Rejected(() => f.Target(User + 15)), "extra plausible member cannot silently inflate the validated active set");
            f = new Fixture(); f.M.Bytes.Remove(f.Users[0] + 0x20);
            Check(Rejected(() => f.Target(User + 1)), "unreadable user ID is not treated as a real ID zero entry");
            f = new Fixture();
            Check(new Fixture(1, "긴 닉네임 테스트 😀").Target().Name == "긴 닉네임 테스트 😀", "long UTF-8 nickname is read without OCR");
            Check(Rejected(() => f.Target(User, Chat + 1)), "another chat ID is rejected");
            Check(Rejected(() => f.Target(0)) && Rejected(() => f.Target(User, 0)), "unknown user or chat ID is rejected");
            using (var reader = f.Reader()) Check(reader.IsParticipantList(f.Window) && !reader.IsParticipantList(new IntPtr(123)), "actual participant control ownership is verified");
            f = new Fixture(); f.M.Type(f.Dialog, ".?AVCOpenChatKickedMemberDlg@@");
            Check(Rejected(() => f.Target()), "banned-member dialog is not mistaken for participant list");
            f = new Fixture(); f.M.U(f.Control + 0x40, 999);
            Check(Rejected(() => f.Target()), "foreign window handle is rejected");
            f = new Fixture(); f.M.U(f.Data[0] + 0x10, f.Items[1]);
            Check(Rejected(() => f.Target()), "item back-reference is checked");
            f = new Fixture(); f.M.U(f.Data[0] + 0x18, f.Control + 8);
            Check(Rejected(() => f.Target()), "control back-reference is checked");
            f = new Fixture(); f.M.U(f.Users[1] + 0x20, (ulong)User);
            Check(Rejected(() => f.Target()), "duplicate user IDs are rejected");
            f = new Fixture(); f.M.U(f.Nodes[1] + 16, f.Nodes[0]);
            Check(Rejected(() => f.Target()), "cyclic tree is rejected without an unbounded traversal");
            f = new Fixture(); f.M.U(f.Control + 0x370, 3);
            Check(Rejected(() => f.Target()), "incomplete tree cannot return an early target");
            f = new Fixture(); f.M.Type(f.Avatars[0], ".?AVTextWidget@xp@@");
            Check(Rejected(() => f.Target()), "target must reference a real profile button widget");
            f = new Fixture(); f.M.Bytes[f.Users[0] + 0x38] = 255;
            Check(Rejected(() => f.Target()), "corrupt nickname data is rejected");
            f = new Fixture(); int sizeReads = 0;
            using (var reader = f.Reader((address, count) => { if (address == f.Control + 0x370 && ++sizeReads == 2) return BitConverter.GetBytes(3UL); return f.M.Read(address, count); }))
                Check(Rejected(() => reader.ReadTarget(f.Window, Chat, User)), "refresh during a read cannot publish a stale target");
            f = new Fixture(40); target = f.Target(User + 38);
            Check(target.UserId == User + 38 && target.AvatarBounds.Top > viewport.Bottom, "offscreen target ID and scroll direction are available without text search");
            Check(ModerationParticipantSearch.Target(target, target.UserId, "ㅇㅇ", viewport).IsEmpty, "offscreen target cannot be clicked before scrolling");
            target = new Fixture().Target(); target.AvatarBounds = new Rectangle(18, -5, 40, 40);
            Check(ModerationParticipantSearch.Target(target, User, "ㅇㅇ", viewport).IsEmpty, "partially clipped avatar cannot be clicked");
            target = new Fixture().Target(); target.NameBounds = new Rectangle(70, 250, 180, 17);
            Check(ModerationParticipantSearch.Target(target, User, "ㅇㅇ", viewport).IsEmpty, "name and avatar from different rows cannot be combined");
            f = new Fixture(92); f.M.I(f.Scroll + 0x114, 450); target = f.Target(User + 15);
            Check(target.ScrollY == 450 && target.ViewportHeight == 280 && target.AvatarBounds.Top == 786,
                "five manual wheel notches read as 450 pixels without changing content coordinates");
            Check(ModerationParticipantSearch.Avatar(target, viewport) == new Rectangle(19, 389, 40, 40),
                "content position is converted using both list origin and current scroll offset");
            Check(ModerationParticipantSearch.Target(target, target.UserId, "ㅇㅇ", viewport).IsEmpty,
                "target below the viewport is still not clickable after a partial scroll");
            f.M.Rect(f.Avatars[15] + 0x80, new Rectangle(18, 208, 40, 40));
            f.M.Rect(f.Labels[15] + 0x80, new Rectangle(70, 220, 180, 17)); f.M.I(f.Scroll + 0x114, 180);
            target = f.Target(User + 15);
            Check(target.AvatarBounds.Top == 786 && ModerationParticipantSearch.Avatar(target, viewport).Top == 659 &&
                ModerationParticipantSearch.Target(target, target.UserId, "ㅇㅇ", viewport).IsEmpty,
                "live wrong-profile regression: stale widget Y=208 cannot select row 4 when the requested user is row 15");
            f.M.I(f.Scroll + 0x114, 630); target = f.Target(User + 15);
            Check(ModerationParticipantSearch.Target(target, target.UserId, "ㅇㅇ", viewport) == new Rectangle(19, 209, 40, 40),
                "scrolled target uses the actual visible profile position");
            Check(ModerationParticipantSearch.Target(target, target.UserId, "ㅇㅇ", new Rectangle(1, 53, 268, 300)).IsEmpty,
                "resized viewport cannot reuse old scroll geometry");
            f.M.I(f.Scroll + 0x114, 0); int steps = 0, position = 0;
            while (steps < 12)
            {
                target = f.Target(User + 15);
                if (!ModerationParticipantSearch.Target(target, target.UserId, "ㅇㅇ", viewport).IsEmpty) break;
                var avatar = ModerationParticipantSearch.Avatar(target, viewport);
                int distance = ModerationParticipantSearch.ScrollStep(avatar, viewport) / 120 * 90;
                position += avatar.Top < viewport.Top ? -distance : distance;
                f.M.I(f.Scroll + 0x114, position); steps++;
            }
            Check(steps == 5 && position == 630, "adaptive wheel steps reach the target instead of bouncing between 540 and 810 pixels");
            f.M.I(f.Scroll + 0x114, 7000);
            Check(Rejected(() => f.Target(User + 15)), "out-of-range native scroll offset is rejected");
            f = new Fixture(10); f.M.Type(f.Scroll, ".?AVUnknownScroll@@");
            Check(Rejected(() => f.Target()), "unrecognized scrollbar cannot supply click coordinates");
            f = new Fixture(10); int positionReads = 0;
            using (var reader = f.Reader((address, count) => { if (address == f.Scroll + 0x114 && ++positionReads == 2) return BitConverter.GetBytes(90); return f.M.Read(address, count); }))
                Check(Rejected(() => reader.ReadTarget(f.Window, Chat, User)), "scroll change during a read invalidates the snapshot");
            f = new Fixture(4); f.ActiveOrder(f.Items[3], f.Items[1], f.Items[0], f.Items[2]);
            Check(f.Target().AvatarBounds.Top == 110, "display tree order is independent of cache traversal and stale widget position");
            f = new Fixture(); f.M.U(f.Items[1] + 0x30, f.Items[0] + 8);
            Check(Rejected(() => f.Target()), "cyclic active display tree is rejected");
            f = new Fixture(); f.M.U(f.Items[1] + 0x20, 0);
            Check(Rejected(() => f.Target()), "active row parent back-reference is verified");
            f = new Fixture(); f.M.I(f.Items[1] + 0x18, 51);
            Check(Rejected(() => f.Target()), "sum of actual row heights must equal the native content extent");
            f = new Fixture(); f.M.U(f.Items[0] + 0x30, 0);
            Check(Rejected(() => f.Target()), "all validated members must belong to the displayed tree");
            f = new Fixture(3); f.M.I(f.Items[0] + 0x18, 60); f.M.I(f.Control + 0xF4, 164);
            Check(f.Target(User + 1).AvatarBounds.Top == 66, "row position sums measured heights instead of multiplying a fixed row index");
            f = new Fixture(); f.M.Type(f.Items[0] + 8, ".?AVUnknown@@", 8);
            Check(Rejected(() => f.Target()), "unsupported base-class pointer cannot determine display order");
            f = new Fixture(); int rowReads = 0;
            using (var reader = f.Reader((address, count) =>
            {
                var bytes = f.M.Read(address, count);
                if (address == f.Items[0] + 0x18 && count == 32 && ++rowReads == 2) bytes[0]++;
                return bytes;
            })) Check(Rejected(() => reader.ReadTarget(f.Window, Chat, User)), "same-count layout change during reading invalidates the target coordinates");
            f = new Fixture(); f.OpenProfile(); f.VerifyProfile();
            Check(true, "opened profile binds the exact window, source list, room, view, controller and user ID");
            f.OpenProfile(1);
            Check(Rejected(() => f.VerifyProfile()), "another opened profile is blocked even with an identical nickname");
            f.OpenProfile(); f.M.U(f.ProfileModel + 0x1B8, (ulong)(Chat + 1));
            Check(Rejected(() => f.VerifyProfile()), "profile model from another room is blocked");
            f.OpenProfile(); f.M.U(f.Profile + 0x40, 1234);
            Check(Rejected(() => f.VerifyProfile()), "profile object for another HWND is blocked");
            f.OpenProfile(); f.M.U(f.Profile + 0x5E0, f.Control + 8);
            Check(Rejected(() => f.VerifyProfile()), "profile from a different participant list is blocked");
            f.OpenProfile(); f.M.U(f.ProfileController + 0x28, f.ProfileModel + 8);
            Check(Rejected(() => f.VerifyProfile()), "controller must use the displayed profile model");
            f.OpenProfile(); f.M.U(f.ProfileController + 0x78, f.Profile + 8);
            Check(Rejected(() => f.VerifyProfile()), "controller must be attached to this displayed view");
            f.OpenProfile(); f.M.Type(f.ProfileModel, ".?AVUnknownProfile@@");
            Check(Rejected(() => f.VerifyProfile()), "unsupported profile model cannot authorize a removal");
            f = new Fixture(); Check(Rejected(() => f.VerifyProfile()), "closed or missing profile blocks removal");
            f.OpenProfile(); Check(Rejected(() => f.VerifyProfile(0)) && Rejected(() => f.VerifyProfile(User, 0)), "invalid expected profile identifiers fail closed");
            f.OpenProfile(); int profileUserReads = 0;
            using (var reader = f.Reader((address, count) => address == f.Users[0] + 0x20 && ++profileUserReads == 2 ? BitConverter.GetBytes((ulong)(User + 1)) : f.M.Read(address, count)))
                Check(Rejected(() => reader.RequireProfile(f.ProfileWindow, f.Window, Chat, User)), "profile target changed during verification is blocked");
            f.OpenProfile(); f.VerifyProfile(); f.M.U(f.ProfileModel + 0x1C8, f.Users[1]);
            Check(Rejected(() => f.VerifyProfile()), "previous successful verification is not reused after the selected profile changes");
            Console.WriteLine("Participant ID checks passed: " + passed + " (no OCR, app input, or removal)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
