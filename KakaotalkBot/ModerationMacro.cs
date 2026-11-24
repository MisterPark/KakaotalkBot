using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
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
        private readonly ModerationMenuImages menuImages;
        private readonly long chatId;
        private bool banListOpen;
        private string banListReadError;
        internal ModerationMacro(string title, int pid, long chatId, Func<bool> allowed)
        {
            if (!StaInputWorker.Instance.IsCurrent) throw new InvalidOperationException("관리 매크로는 STA 명령에서만 실행합니다.");
            this.allowed = allowed; this.chatId = chatId;
            var rooms = Desktop.Windows().Where(w => w.Process == pid && w.Title == title && w.Visible).ToArray();
            if (rooms.Length != 1) throw new InvalidOperationException("대상 채팅방 창을 유일하게 찾지 못했습니다.");
            room = rooms[0]; assets = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ModerationAssets");
            menuImages = new ModerationMenuImages(Path.Combine(assets, "MenuCache"), assets);
            vision = new Vision(assets);
            Guard();
            if (!Desktop.SetForegroundWindow(room.Handle)) throw new InvalidOperationException("대상 방을 활성화하지 못했습니다.");
        }
        private void Guard() { if (!allowed()) throw new InvalidOperationException("방·권한·수신 세션이 변경되어 매크로를 중단했습니다."); }
        private bool Related(Window w, List<Window> all)
        {
            return ModerationMenu.Related(w, room, all, ModerationMenu.ActiveOwner());
        }
        private void Click(Window w, Rectangle bounds, ActionKind kind = ActionKind.Click, int wheelDelta = 360, Action verify = null)
        {
            Guard();
            using (var frame = Desktop.Capture(w.Handle))
            {
                if (frame == null || frame.Bounds != w.Bounds || !Desktop.Apply(frame, new Decision(kind, "관리 명령", new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2)) { WheelDelta = wheelDelta },
                    () => { if (!allowed()) return true; if (verify != null) verify(); return false; }, true))
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
                // 내보내기 팝업은 180x100보다 작을 수 있습니다. 자식/가려진 부모 캡처를 메뉴로 쓰지 않습니다.
                foreach (var w in all.Where(w => Desktop.IsCaptureWindow(w) && Related(w, all) && w.Bounds.Width >= 16 && w.Bounds.Height >= 16).OrderBy(w => w.Bounds.Width * w.Bounds.Height))
                    using (var frame = Desktop.Capture(w.Handle))
                        if (frame != null && predicate(w, frame.Image)) return w;
                Thread.Sleep(30);
            }
            throw new TimeoutException(step + " 화면을 찾지 못했습니다.");
        }
        private void TextClick(string text)
        {
            var timer = Stopwatch.StartNew(); var observations = new HashSet<string>();
            var ocrFrames = new Dictionary<IntPtr, ulong>();
            while (timer.ElapsedMilliseconds < 5000)
            {
                Guard();
                var all = Desktop.Windows(false, true); var menuOwner = ModerationMenu.ActiveOwner();
                var candidates = all.Where(w => ModerationMenu.TextCandidate(w, room, all, menuOwner))
                    .OrderBy(w => w.ClassName == "#32768" ? 0 : 1).ThenBy(w => w.Bounds.Width * w.Bounds.Height).ToArray();
                foreach (var window in candidates)
                {
                    Rectangle bounds; string observed;
                    if (ModerationMenu.TryNativeBounds(window, text, out bounds, out observed)) { Click(window, bounds); return; }
                    using (var frame = Desktop.Capture(window.Handle))
                    {
                        if (frame == null) continue;
                        var pixels = new Pixels(frame.Image);
                        if (window.ClassName == "EVA_Menu" && menuImages.FindIcon(text, pixels, out bounds)) { Click(window, bounds); return; }
                        if (menuImages.Find(text, pixels, out bounds)) { Click(window, bounds); return; }
                        ulong fingerprint = pixels.Fingerprint(new Rectangle(Point.Empty, frame.Image.Size)), previous;
                        if (ocrFrames.TryGetValue(window.Handle, out previous) && previous == fingerprint) continue;
                        ocrFrames[window.Handle] = fingerprint;
                        var labels = ModerationOcr.Read(frame.Image);
                        observed = string.Join(" | ", labels.Select(item => item.Text));
                        var label = ModerationMenu.Exact(labels, text);
                        if (label != null) { menuImages.Remember(text, frame.Image, label.Bounds); Click(window, label.Bounds); return; }
                    }
                    observations.Add(window.ClassName + " " + window.Bounds.Width + "x" + window.Bounds.Height + " : " + observed);
                }
                if (candidates.Length == 0) observations.Add("대상 방에 연결된 메뉴 창 없음");
                Thread.Sleep(30);
            }
            try { File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "moderation-menu.log"), DateTimeOffset.Now + " " + text + Environment.NewLine + string.Join(Environment.NewLine, observations)); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            throw new TimeoutException(text + " 메뉴를 찾지 못했습니다. 해당 메뉴가 표시되어 있는지 확인해 주세요.");
        }
        private void Escape(Window window)
        {
            using (var frame = Desktop.Capture(window.Handle))
                if (frame != null) Desktop.Apply(frame, new Decision(ActionKind.Escape, "관리 화면 닫기"), () => false);
        }
        private void OpenBanManagement()
        {
            Window dialog = null; var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < 5000 && dialog == null)
            {
                Guard(); dialog = ModerationNavigation.Dialog(room, Desktop.Windows(false, true));
                if (dialog == null) Thread.Sleep(30);
            }
            if (dialog == null) throw new InvalidOperationException("오픈채팅 관리 창의 스크롤 영역을 찾지 못했습니다.");
            Guard();
            if (!Desktop.SetForegroundWindow(dialog.Handle)) throw new InvalidOperationException("오픈채팅 관리 창을 활성화하지 못했습니다.");
            var observations = new HashSet<string>(); ulong fingerprint = 0;
            Func<Window> content = () =>
            {
                Guard();
                var all = Desktop.Windows(false, true);
                var current = all.SingleOrDefault(w => w.Handle == dialog.Handle && w.Visible && Related(w, all));
                var pane = current == null ? null : ModerationNavigation.Content(current, all);
                if (pane == null) throw new InvalidOperationException("오픈채팅 관리 창이 닫혔거나 스크롤 영역이 바뀌었습니다.");
                return pane;
            };
            Func<bool> selectVisible = () =>
            {
                var pane = content();
                using (var frame = Desktop.Capture(pane.Handle))
                {
                    if (frame == null) throw new InvalidOperationException("관리 창의 내용을 읽지 못했습니다.");
                    var pixels = new Pixels(frame.Image); Rectangle button;
                    fingerprint = pixels.Fingerprint(new Rectangle(Point.Empty, frame.Image.Size));
                    observations.Add(pane.ClassName + " " + frame.Image.Width + "x" + frame.Image.Height + " 화면=" + fingerprint);
                    if (!menuImages.FindButton("내보낸 상대 관리", pixels, out button)) return false;
                    Guard(); Click(pane, button); return true;
                }
            };
            Action scroll = () =>
            {
                var pane = content(); var point = ModerationNavigation.ScrollPoint(pane.Bounds.Size);
                Click(pane, new Rectangle(point.X - 1, point.Y - 1, 2, 2), ActionKind.ScrollDown);
            };
            Func<bool> changed = () =>
            {
                var settle = Stopwatch.StartNew();
                while (settle.ElapsedMilliseconds < 750)
                {
                    var pane = content();
                    using (var frame = Desktop.Capture(pane.Handle))
                    {
                        if (frame == null) throw new InvalidOperationException("관리 창이 닫혔습니다.");
                        if (new Pixels(frame.Image).Fingerprint(new Rectangle(Point.Empty, frame.Image.Size)) != fingerprint)
                        { Thread.Sleep(100); return true; }
                    }
                    Thread.Sleep(30);
                }
                return false;
            };
            if (ModerationNavigation.Search(Guard, selectVisible, scroll, changed)) return;
            try { File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "moderation-menu.log"), DateTimeOffset.Now + " 내보낸 상대 관리 (스크롤 탐색)" + Environment.NewLine + string.Join(Environment.NewLine, observations)); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            throw new InvalidOperationException("관리 창을 아래로 탐색했지만 내보낸 상대 관리 버튼을 찾지 못했습니다.");
        }
        private Window OpenParticipantProfile(long userId, string nickname)
        {
            Rectangle button = Rectangle.Empty;
            var all = Desktop.Windows(false, true);
            if (!all.Any(w => Desktop.IsParticipantPopup(w, room, all)))
            {
                var window = Wait((w, image) => { if (w.Handle != room.Handle) return false; button = vision.Read(image, false, Stage.OpenList).Members; return !button.IsEmpty; }, "참여자 목록 버튼");
                Click(window, button);
            }
            var timer = Stopwatch.StartNew(); string failure = "참여자 목록을 준비 중입니다.";
            using (var reader = new ParticipantMemory((int)room.Process))
            while (timer.ElapsedMilliseconds < 10000)
            {
                Guard(); all = Desktop.Windows(false, true);
                var popups = all.Where(w => Desktop.IsParticipantPopup(w, room, all)).ToArray();
                if (popups.Length > 1) throw new InvalidOperationException("대상 방의 참여자 목록 창이 여러 개입니다.");
                if (popups.Length == 1)
                {
                    var popup = popups[0];
                    var lists = all.Where(w => w.Parent == popup.Handle && w.ClassName == "EVA_VH_ListControl_Dblclk" &&
                        w.Process == room.Process && reader.IsParticipantList(w.Handle)).ToArray();
                    if (lists.Length > 1) throw new InvalidOperationException("참여자 목록 객체가 여러 개입니다.");
                    if (lists.Length == 1)
                    {
                        var list = lists[0]; ParticipantMember target = null;
                        try { target = reader.ReadTarget(list.Handle, chatId, userId); if (target == null) failure = "대상 사용자 ID를 목록에서 확인하지 못했습니다."; }
                        catch (InvalidOperationException error) { failure = error.Message; }
                        if (target != null)
                        {
                            if (target.Name != nickname) throw new InvalidOperationException("대상 참여자의 닉네임이 변경되었습니다. 다시 명령해 주세요.");
                            var viewport = new Rectangle(list.Bounds.Left - popup.Bounds.Left, list.Bounds.Top - popup.Bounds.Top, list.Bounds.Width, list.Bounds.Height);
                            var avatar = ModerationParticipantSearch.Target(target, userId, nickname, viewport);
                            if (!avatar.IsEmpty)
                            {
                                using (var frame = Desktop.Capture(popup.Handle))
                                {
                                    if (frame == null) throw new InvalidOperationException("참여자 목록 창이 닫혔습니다.");
                                    var currentList = Desktop.Windows(false, true).SingleOrDefault(w => w.Handle == list.Handle && w.Parent == popup.Handle && w.Process == room.Process);
                                    if (frame.Bounds != popup.Bounds || currentList == null || currentList.Bounds != list.Bounds) { Thread.Sleep(50); continue; }
                                    ParticipantMember fresh;
                                    try { fresh = reader.ReadTarget(list.Handle, chatId, userId); }
                                    catch (InvalidOperationException error) { failure = error.Message; Thread.Sleep(50); continue; }
                                    if (ModerationParticipantSearch.Target(fresh, userId, nickname, viewport) != avatar) { Thread.Sleep(50); continue; }
                                    Guard();
                                    if (!Desktop.Apply(frame, new Decision(ActionKind.Click, "사용자 ID로 확인한 참여자", new Point(avatar.Left + avatar.Width / 2, avatar.Top + avatar.Height / 2)), () => !allowed()))
                                        throw new InvalidOperationException("참여자 창이 이동했거나 가려져 클릭하지 않았습니다.");
                                }
                                return list;
                            }
                            // 화면 밖의 행은 위젯 좌표로 방향을 결정하고 실제 목록 영역에서 휠만 움직입니다.
                            var position = ModerationParticipantSearch.Avatar(target, viewport);
                            if (position.Width >= 12 && position.Height >= 12 &&
                                (position.Top < viewport.Top || position.Bottom > viewport.Bottom))
                            {
                                var direction = position.Top < viewport.Top ? ActionKind.ScrollUp : ActionKind.ScrollDown;
                                Click(list, new Rectangle(Math.Max(1, list.Bounds.Width - 25), list.Bounds.Height / 2, 2, 2), direction, ModerationParticipantSearch.ScrollStep(position, viewport));
                                Thread.Sleep(100);
                            }
                            failure = "대상 참여자의 프로필 버튼이 아직 목록 안에 완전히 표시되지 않았습니다.";
                        }
                    }
                }
                Thread.Sleep(50);
            }
            throw new TimeoutException("대상 참여자 확인 실패: " + failure);
        }
        internal void Kick(long userId, string nickname)
        {
            var list = OpenParticipantProfile(userId, nickname);
            using (var reader = new ParticipantMemory((int)room.Process))
            {
                Rectangle button = Rectangle.Empty;
                var profile = Wait((w, image) =>
                {
                    if (w.Owner != list.Parent || !Desktop.IsProfileWindow(w, room, Desktop.Windows(false, true))) return false;
                    reader.RequireProfile(w.Handle, list.Handle, chatId, userId);
                    button = vision.FindProfileAction(image); return !button.IsEmpty && Desktop.IsExposed(w, button);
                }, "프로필 내보내기 메뉴");
                Action verify = () =>
                {
                    Guard();
                    var all = Desktop.Windows(false, true);
                    var current = all.SingleOrDefault(w => w.Handle == profile.Handle && w.Owner == list.Parent);
                    if (current == null || !Desktop.IsProfileWindow(current, room, all))
                        throw new InvalidOperationException("확인한 참여자 프로필이 닫히거나 바뀌었습니다.");
                    reader.RequireProfile(current.Handle, list.Handle, chatId, userId);
                };
                Click(profile, button, verify: verify);
                var window = Wait((w, image) =>
                {
                    verify();
                    if (!ModerationMenu.Related(w, profile, Desktop.Windows(false, true), ModerationMenu.ActiveOwner())) return false;
                    button = vision.Read(image, false, Stage.RemoveMenu).Remove; return !button.IsEmpty && Desktop.IsExposed(w, button);
                }, "내보내기");
                Click(window, button, verify: verify);
                window = Wait((w, image) =>
                {
                    verify();
                    if (!ModerationMenu.Related(w, profile, Desktop.Windows(false, true), ModerationMenu.ActiveOwner())) return false;
                    button = vision.Read(image, false, Stage.Confirm).Confirm; return !button.IsEmpty && Desktop.IsExposed(w, button);
                }, "내보내기 확인");
                Click(window, button, ActionKind.ConfirmRemoval, verify: verify);
            }
            // 최종 클릭까지를 매크로 완료로 기록하며 실제 퇴장 여부를 기다리지 않습니다.
        }
        private Window CurrentBanList(out List<BannedMember> members)
        {
            var all = Desktop.Windows(false, true); members = null;
            banListOpen = false; banListReadError = null;
            using (var reader = new BannedMemberMemory((int)room.Process))
                foreach (var window in all.Where(w => Related(w, all) && w.ClassName == "EVA_VH_ListControl_Dblclk"))
                {
                    if (!reader.IsBanListWindow(window.Handle)) continue;
                    banListOpen = true;
                    try { members = reader.ReadList(window.Handle, chatId); return window; }
                    catch (InvalidOperationException error) { banListReadError = error.Message; }
                }
            return null;
        }
        private Window OpenBanList(out List<BannedMember> members)
        {
            var current = CurrentBanList(out members);
            if (!banListOpen)
            {
                if (ModerationNavigation.Dialog(room, Desktop.Windows(false, true)) == null)
                {
                    using (var frame = Desktop.Capture(room.Handle))
                    using (var template = new Bitmap(Path.Combine(assets, "menu.png")))
                    {
                        Rectangle button;
                        if (frame == null || !new Pixels(frame.Image).Find(new Pixels(template), new Rectangle(0, 0, frame.Image.Width, Math.Min(100, frame.Image.Height)), out button))
                            throw new InvalidOperationException("채팅방 메뉴 버튼을 찾지 못했습니다.");
                        Click(room, button);
                    }
                    TextClick("오픈채팅 관리");
                }
                OpenBanManagement();
            }
            var timer = Stopwatch.StartNew();
            try
            {
                members = BannedListLoading.Wait(() =>
                {
                    List<BannedMember> snapshot;
                    current = CurrentBanList(out snapshot);
                    return current == null ? null : snapshot;
                }, Guard, () => timer.ElapsedMilliseconds, Thread.Sleep);
                return current;
            }
            catch (TimeoutException)
            {
                if (banListReadError != null) throw new InvalidOperationException("내보낸 목록 읽기 실패: " + banListReadError);
                if (!banListOpen) throw new InvalidOperationException("내보낸 상대 목록 창을 확인하지 못했습니다.");
                throw;
            }
        }
        internal void CloseOwnedDialogs()
        {
            for (int i = 0; i < 4; i++)
            {
                var all = Desktop.Windows(false, true);
                IntPtr handle = Desktop.GetForegroundWindow();
                var foreground = all.FirstOrDefault(w => w.Handle == handle);
                if (foreground == null || !Related(foreground, all)) return;
                if (foreground.Handle == room.Handle && !all.Any(w => w.ClassName == "#32768" && Related(w, all))) return;
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
                    var target = reader.ReadList(list.Handle, chatId).SingleOrDefault(r => r.UserId == userId && r.KickedAt == kickedAt);
                    if (target == null) throw new InvalidOperationException("대상 목록이 변경되었습니다.");
                    using (var frame = Desktop.Capture(list.Handle))
                    {
                        if (frame == null) throw new InvalidOperationException("관리 목록이 닫혔습니다.");
                        if (target.NameBounds.Top >= 0 && target.NameBounds.Bottom <= frame.Image.Height)
                        {
                            var button = ReleaseButton(ModerationOcr.Read(frame.Image), target.NameBounds);
                            if (!button.IsEmpty)
                            {
                                var fresh = reader.ReadList(list.Handle, chatId).Single(r => r.UserId == userId && r.KickedAt == kickedAt);
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
