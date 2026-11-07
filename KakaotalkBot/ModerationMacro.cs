using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using OpenChatRoomCleaner;

namespace KakaotalkBot
{
    // 한 호출 전체가 STA 명령 한 건입니다. 클릭 이후 서버 성공 여부를 재조회하지 않습니다.
    internal sealed class ModerationMacro
    {
        private readonly Window room;
        private readonly Func<bool> allowed;
        private readonly Vision vision;
        private readonly string assets;
        private readonly long linkId;
        internal ModerationMacro(string title, int pid, long linkId, Func<bool> allowed)
        {
            if (!StaInputWorker.Instance.IsCurrent) throw new InvalidOperationException("관리 매크로는 STA 명령에서만 실행합니다.");
            this.allowed = allowed; this.linkId = linkId;
            var rooms = Desktop.Windows().Where(w => w.Process == pid && w.Title == title && w.Visible).ToArray();
            if (rooms.Length != 1) throw new InvalidOperationException("대상 채팅방 창을 유일하게 찾지 못했습니다.");
            room = rooms[0]; assets = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ModerationAssets");
            vision = new Vision(assets);
            Guard();
            if (!Desktop.SetForegroundWindow(room.Handle)) throw new InvalidOperationException("대상 방을 활성화하지 못했습니다.");
        }
        private void Guard() { if (!allowed()) throw new InvalidOperationException("방·권한·수신 세션이 변경되어 매크로를 중단했습니다."); }
        private bool Related(Window w, List<Window> all)
        {
            if (w.Process != room.Process || !w.Visible) return false;
            var seen = new HashSet<IntPtr>();
            while (w != null && seen.Add(w.Handle))
            {
                if (w.Handle == room.Handle) return true;
                IntPtr ancestor = w.Owner != IntPtr.Zero ? w.Owner : w.Parent;
                w = all.FirstOrDefault(x => x.Handle == ancestor);
            }
            return false;
        }
        private void Click(Window w, Rectangle bounds, ActionKind kind = ActionKind.Click)
        {
            Guard();
            using (var frame = Desktop.Capture(w.Handle))
            {
                if (frame == null || !Desktop.Apply(frame, new Decision(kind, "관리 명령", new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2)), () => !allowed(), true))
                    throw new InvalidOperationException("관리 창이 이동했거나 다른 창에 가려 클릭하지 않았습니다.");
            }
        }
        private Window Wait(Func<Window, Bitmap, bool> predicate, string step)
        {
            var timer = Stopwatch.StartNew();
            while (timer.ElapsedMilliseconds < 5000)
            {
                Guard();
                var all = Desktop.Windows(false, true);
                foreach (var w in all.Where(w => Related(w, all) && w.Bounds.Width >= 180 && w.Bounds.Height >= 100).OrderBy(w => w.Bounds.Width * w.Bounds.Height))
                    using (var frame = Desktop.Capture(w.Handle))
                        if (frame != null && predicate(w, frame.Image)) return w;
                Thread.Sleep(30);
            }
            throw new TimeoutException(step + " 화면을 찾지 못했습니다.");
        }
        private static OcrLabel Exact(IEnumerable<OcrLabel> labels, string text)
        {
            string expected = ModerationOcr.Normalize(text);
            var matches = labels.Where(l => ModerationOcr.Normalize(l.Text) == expected).ToArray();
            return matches.Length == 1 ? matches[0] : null;
        }
        private void TextClick(string text)
        {
            Rectangle bounds = Rectangle.Empty;
            var window = Wait((w, image) => { var label = Exact(ModerationOcr.Read(image), text); if (label == null) return false; bounds = label.Bounds; return true; }, text);
            Click(window, bounds);
        }
        private void Escape(Window window)
        {
            using (var frame = Desktop.Capture(window.Handle))
                if (frame != null) Desktop.Apply(frame, new Decision(ActionKind.Escape, "관리 화면 닫기"), () => false);
        }
        internal void Kick(string nickname)
        {
            Rectangle button = Rectangle.Empty;
            var window = Wait((w, image) => { if (w.Handle != room.Handle) return false; button = vision.Read(image, false, Stage.OpenList).Members; return !button.IsEmpty; }, "참여자 목록 버튼");
            Click(window, button);
            Window popup = null;
            window = Wait((w, image) =>
            {
                if (!Desktop.IsParticipantPopup(w, room, Desktop.Windows(false, true))) return false;
                using (var search = new Bitmap(Path.Combine(assets, "list-search.png")))
                    return new Pixels(image).Find(new Pixels(search), new Rectangle(0, 0, image.Width, Math.Min(80, image.Height)), out button);
            }, "참여자 검색");
            popup = window;
            Click(window, button);
            SendKeys.SendWait("^a"); Clipboard.SetText(nickname); SendKeys.SendWait("^v");
            Rectangle avatar = Rectangle.Empty;
            window = Wait((w, image) =>
            {
                if (w.Handle != popup.Handle) return false;
                var observation = vision.Read(image);
                var name = Exact(ModerationOcr.Read(image), nickname);
                if (name == null || !observation.ListValid) return false;
                var rows = observation.Rows.Where(r => !r.Protected && Math.Abs(r.Click.Y - (name.Bounds.Top + name.Bounds.Height / 2)) < 25).ToArray();
                if (rows.Length != 1) return false;
                avatar = new Rectangle(rows[0].Click.X - 2, rows[0].Click.Y - 2, 4, 4); return true;
            }, "대상 참여자");
            Click(window, avatar);
            window = Wait((w, image) =>
            {
                if (!Desktop.IsProfileWindow(w, room, Desktop.Windows(false, true))) return false;
                button = vision.FindProfileAction(image); return !button.IsEmpty;
            }, "프로필 내보내기 메뉴");
            Click(window, button);
            window = Wait((w, image) => { button = vision.Read(image, false, Stage.RemoveMenu).Remove; return !button.IsEmpty; }, "내보내기");
            Click(window, button);
            window = Wait((w, image) => { button = vision.Read(image, false, Stage.Confirm).Confirm; return !button.IsEmpty; }, "내보내기 확인");
            Click(window, button, ActionKind.ConfirmRemoval);
            // 최종 클릭까지를 매크로 완료로 기록하며 실제 퇴장 여부를 기다리지 않습니다.
        }
        private Window CurrentBanList(out List<BannedMember> members)
        {
            var all = Desktop.Windows(false, true); members = null;
            using (var reader = new BannedMemberMemory((int)room.Process))
                foreach (var window in all.Where(w => Related(w, all) && w.ClassName == "EVA_VH_ListControl_Dblclk"))
                {
                    try { members = reader.ReadList(window.Handle, linkId); return window; }
                    catch (InvalidOperationException) { }
                }
            return null;
        }
        private Window OpenBanList(out List<BannedMember> members)
        {
            var current = CurrentBanList(out members);
            if (current != null) return current;
            using (var frame = Desktop.Capture(room.Handle))
            using (var template = new Bitmap(Path.Combine(assets, "menu.png")))
            {
                Rectangle button;
                if (frame == null || !new Pixels(frame.Image).Find(new Pixels(template), new Rectangle(0, 0, frame.Image.Width, Math.Min(100, frame.Image.Height)), out button))
                    throw new InvalidOperationException("채팅방 메뉴 버튼을 찾지 못했습니다.");
                Click(room, button);
            }
            TextClick("오픈채팅 관리");
            TextClick("내보낸 상대 관리");
            var timer = Stopwatch.StartNew();
            while (timer.ElapsedMilliseconds < 5000)
            {
                Guard(); current = CurrentBanList(out members);
                if (current != null) return current;
                Thread.Sleep(30);
            }
            throw new InvalidOperationException("내보낸 목록이 비었거나 현재 버전의 목록 객체를 확인하지 못했습니다.");
        }
        internal void CloseOwnedDialogs()
        {
            for (int i = 0; i < 4; i++)
            {
                var all = Desktop.Windows(false, true);
                IntPtr handle = Desktop.GetForegroundWindow();
                var foreground = all.FirstOrDefault(w => w.Handle == handle);
                if (foreground == null || foreground.Handle == room.Handle || !Related(foreground, all)) return;
                Escape(foreground);
                var wait = Stopwatch.StartNew();
                while (Desktop.GetForegroundWindow() == handle && wait.ElapsedMilliseconds < 500) Thread.Sleep(20);
            }
        }
        internal List<BannedMember> ReadBans()
        {
            List<BannedMember> rows; var window = OpenBanList(out rows);
            Escape(window); return rows;
        }
        internal static Rectangle ReleaseButton(IEnumerable<OcrLabel> labels, Rectangle nameBounds)
        {
            var buttons = labels.Where(l => ModerationOcr.Normalize(l.Text) == "해제" && l.Bounds.Left > nameBounds.Left && Math.Abs(l.Bounds.Top - nameBounds.Top) < 20).ToArray();
            return buttons.Length == 1 ? buttons[0].Bounds : Rectangle.Empty;
        }
        internal void Release(long userId, long kickedAt)
        {
            List<BannedMember> rows; var list = OpenBanList(out rows);
            if (!rows.Any(r => r.UserId == userId && r.KickedAt == kickedAt)) throw new InvalidOperationException("지정한 목록 기록이 없습니다. /내보낸목록으로 다시 확인하세요.");
            using (var reader = new BannedMemberMemory((int)room.Process))
            {
                for (int page = 0; page < 100; page++)
                {
                    Guard();
                    var target = reader.ReadList(list.Handle, linkId).SingleOrDefault(r => r.UserId == userId && r.KickedAt == kickedAt);
                    if (target == null) throw new InvalidOperationException("대상 목록이 변경되었습니다.");
                    using (var frame = Desktop.Capture(list.Handle))
                    {
                        if (frame == null) throw new InvalidOperationException("관리 목록이 닫혔습니다.");
                        if (target.NameBounds.Top >= 0 && target.NameBounds.Bottom <= frame.Image.Height)
                        {
                            var button = ReleaseButton(ModerationOcr.Read(frame.Image), target.NameBounds);
                            if (!button.IsEmpty)
                            {
                                var fresh = reader.ReadList(list.Handle, linkId).Single(r => r.UserId == userId && r.KickedAt == kickedAt);
                                if (fresh.NameBounds != target.NameBounds) continue;
                                Click(list, button); Escape(list); return;
                            }
                        }
                        Click(list, new Rectangle(frame.Image.Width / 2, frame.Image.Height / 2, 2, 2), target.NameBounds.Top < 0 ? ActionKind.ScrollUp : ActionKind.ScrollDown);
                    }
                    Thread.Sleep(50);
                }
            }
            Escape(list);
            throw new InvalidOperationException("대상 행의 해제 버튼을 찾지 못했습니다.");
        }
    }
}
