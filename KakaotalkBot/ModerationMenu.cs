using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using OpenChatRoomCleaner;

namespace KakaotalkBot
{
    internal sealed class ModerationMenuItem
    {
        internal int Index;
        internal string Text;
        internal bool Enabled;
    }

    internal static class ModerationMenu
    {
        // 표준 팝업 메뉴는 Owner/Parent가 0일 수 있습니다. 현재 메뉴 소유 창까지 검증합니다.
        internal static IntPtr ActiveOwner()
        {
            var info = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            return GetGUIThreadInfo(0, ref info) && (info.Flags & 0x14) != 0 ? info.MenuOwner : IntPtr.Zero;
        }
        private static bool OwnedBy(IntPtr handle, Window room, IList<Window> windows)
        {
            var pending = new Stack<IntPtr>(); var seen = new HashSet<IntPtr>(); pending.Push(handle);
            while (pending.Count > 0)
            {
                var current = pending.Pop();
                if (current == IntPtr.Zero || !seen.Add(current)) continue;
                if (current == room.Handle) return true;
                var window = windows.FirstOrDefault(w => w.Handle == current && w.Process == room.Process);
                if (window == null) continue;
                pending.Push(window.Parent); pending.Push(window.Owner);
            }
            return false;
        }
        internal static bool Related(Window window, Window room, IList<Window> windows, IntPtr menuOwner)
        {
            if (!window.Visible || window.Process != room.Process) return false;
            if (OwnedBy(window.Handle, room, windows)) return true;
            return window.ClassName == "#32768" && window.Thread == room.Thread && OwnedBy(menuOwner, room, windows);
        }
        internal static bool TextCandidate(Window window, Window room, IList<Window> windows, IntPtr menuOwner)
        {
            // 채팅 본문·입력창을 메뉴로 오인하지 않습니다. 메뉴에는 180x100 크기 제한을 적용하지 않습니다.
            return window.Handle != room.Handle && window.Bounds.Width >= 16 && window.Bounds.Height >= 16 &&
                (window.ClassName == "#32768" || window.ClassName.StartsWith("EVA_", StringComparison.Ordinal) &&
                    !window.ClassName.Contains("ListControl")) && Related(window, room, windows, menuOwner);
        }
        internal static string Normalize(string text)
        {
            text = (text ?? "").Split('\t')[0].Replace("&", "").Trim().TrimEnd('>', '›', '»', '▶', '▸');
            return new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
        }
        internal static OcrLabel Exact(IEnumerable<OcrLabel> labels, string text)
        {
            string expected = Normalize(text);
            var found = labels.Where(label => Normalize(label.Text) == expected).ToArray();
            return found.Length == 1 ? found[0] : null;
        }
        internal static List<ModerationMenuItem> ReadItems(IntPtr menu)
        {
            var result = new List<ModerationMenuItem>();
            int count = GetMenuItemCount(menu);
            if (count < 1 || count > 256) return result;
            for (int index = 0; index < count; index++)
            {
                var info = new MenuItemInfo { Size = (uint)Marshal.SizeOf(typeof(MenuItemInfo)), Mask = 0x141 };
                if (!GetMenuItemInfo(menu, (uint)index, true, ref info) || (info.Type & 0x800) != 0 || info.Length < 1 || info.Length > 1024) continue;
                IntPtr buffer = Marshal.AllocHGlobal(checked(((int)info.Length + 1) * 2));
                try
                {
                    info.Text = buffer; info.Length++;
                    if (GetMenuItemInfo(menu, (uint)index, true, ref info))
                        result.Add(new ModerationMenuItem { Index = index, Text = Marshal.PtrToStringUni(buffer, (int)info.Length), Enabled = (info.State & 3) == 0 });
                }
                finally { Marshal.FreeHGlobal(buffer); }
            }
            return result;
        }
        internal static bool TryNativeBounds(Window window, string text, out Rectangle bounds, out string observed)
        {
            bounds = Rectangle.Empty; observed = "";
            if (window.ClassName != "#32768") return false;
            UIntPtr value;
            // MN_GETHMENU: 메뉴 조회만 수행하며 명령 메시지를 보내지 않습니다.
            if (SendMessageTimeout(window.Handle, 0x1E1, UIntPtr.Zero, IntPtr.Zero, 2, 200, out value) == IntPtr.Zero) return false;
            IntPtr menu = (IntPtr)unchecked((long)value.ToUInt64());
            if (menu == IntPtr.Zero) return false;
            var items = ReadItems(menu);
            observed = string.Join(" | ", items.Select(item => item.Text));
            var matches = items.Where(item => Normalize(item.Text) == Normalize(text)).ToArray();
            if (matches.Length > 1) throw new InvalidOperationException("같은 이름의 관리 메뉴가 여러 개여서 선택하지 않았습니다.");
            if (matches.Length != 1) return false;
            if (!matches[0].Enabled) throw new InvalidOperationException("카카오톡에서 '" + text + "' 메뉴가 비활성화되어 있습니다. 봇 계정의 방 관리 권한을 확인해 주세요.");
            NativeRect rect;
            if (!GetMenuItemRect(IntPtr.Zero, menu, (uint)matches[0].Index, out rect)) return false;
            bounds = RelativeBounds(window.Bounds, Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom));
            return !bounds.IsEmpty;
        }
        internal static Rectangle RelativeBounds(Rectangle window, Rectangle item)
        {
            if (item.Width <= 0 || item.Height <= 0 || !window.Contains(item)) return Rectangle.Empty;
            item.Offset(-window.Left, -window.Top); return item;
        }

        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { internal int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct GuiThreadInfo
        { internal int Size, Flags; internal IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret; internal NativeRect CaretRect; }
        [StructLayout(LayoutKind.Sequential)] private struct MenuItemInfo
        {
            internal uint Size, Mask, Type, State, Id;
            internal IntPtr SubMenu, Checked, Unchecked, Data, Text;
            internal uint Length;
            internal IntPtr Bitmap;
        }
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
        [DllImport("user32.dll")] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wparam, IntPtr lparam, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll")] private static extern int GetMenuItemCount(IntPtr menu);
        [DllImport("user32.dll", EntryPoint = "GetMenuItemInfoW")] private static extern bool GetMenuItemInfo(IntPtr menu, uint index, bool byPosition, ref MenuItemInfo info);
        [DllImport("user32.dll")] private static extern bool GetMenuItemRect(IntPtr window, IntPtr menu, uint index, out NativeRect rect);
    }
}
