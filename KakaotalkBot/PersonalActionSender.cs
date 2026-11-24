using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using OpenChatRoomCleaner;

namespace KakaotalkBot
{
    internal static class PersonalActionSender
    {
        internal static void RequireNoDraft(ChatRoomInfo target)
        {
            var room = Desktop.Windows().SingleOrDefault(w => w.Process == target.ProcessId && w.Title == target.Name && w.Root == w.Handle && w.Owner == IntPtr.Zero);
            PersonalActionMemory.Require(room != null);
            using (var memory = new PersonalActionMemory(target.ProcessId))
                if (memory.Draft(memory.Room(room.Handle, target.ChatId)) != null)
                    throw new InvalidOperationException("선택해 둔 이모티콘이 있어 자동응답을 멈췄습니다.");
        }
        private sealed class Target
        {
            internal Window Window;
            internal Rectangle Bounds;
        }
        private static Target At(IntPtr window, Rectangle bounds, int pid)
        {
            var w = Desktop.Windows(false, true).SingleOrDefault(x => x.Handle == window && x.Process == pid);
            PersonalActionMemory.Require(w != null); return new Target { Window = w, Bounds = bounds };
        }
        private static bool Visible(Target target)
        { return target.Bounds.Width > 1 && target.Bounds.Height > 1 && new Rectangle(Point.Empty, target.Window.Bounds.Size).Contains(target.Bounds); }
        private static void Act(Func<Target> locate, Action require, bool click)
        {
            require(); var first = locate(); PersonalActionMemory.Require(Visible(first));
            var current = locate();
            PersonalActionMemory.Require(first.Window.Handle == current.Window.Handle && first.Window.Bounds == current.Window.Bounds && first.Bounds == current.Bounds && Desktop.IsExposed(current.Window, current.Bounds));
            require();
            PersonalActionMemory.Require(Desktop.IsForegroundProcess(current.Window.Process) && new[] { 0x10, 0x11, 0x12, 0x01, 0x02 }.All(k => (Desktop.GetAsyncKeyState(k) & 0x8000) == 0));
            var p = new Point(current.Bounds.Left + current.Bounds.Width / 2, current.Bounds.Top + current.Bounds.Height / 2);
            if (!click) { PersonalActionMemory.Require(SetCursorPos(current.Window.Bounds.Left + p.X, current.Window.Bounds.Top + p.Y)); return; }
            using (var frame = new Frame { Handle = current.Window.Handle, Bounds = current.Window.Bounds })
                PersonalActionMemory.Require(Desktop.Apply(frame, new Decision(ActionKind.Click, "개인 응답 항목 선택", p), () => { require(); return false; }));
        }
        private static Window Popup(Window room, PersonalActionMemory memory, string type)
        {
            var candidates = Desktop.Windows().Where(w => w.Process == room.Process && w.Owner == room.Handle && memory.Dialog(w.Handle, type) != 0).ToArray();
            PersonalActionMemory.Require(candidates.Length <= 1); return candidates.SingleOrDefault();
        }
        private static Window WaitPopup(Window room, PersonalActionMemory memory, string type, Action require)
        {
            for (int i = 0; i < 20; i++) { require(); var popup = Popup(room, memory, type); if (popup != null) return popup; Thread.Sleep(50); }
            throw new InvalidOperationException("카카오톡 선택창을 열지 못했습니다.");
        }
        internal static T OpenOrReusePicker<T>(Func<T> find, Action open, Func<T> wait, Action<T> validate) where T : class
        {
            // 같은 방의 열린 선택창은 다시 열기 버튼을 눌러 닫지 않고 재사용합니다.
            var picker = find();
            if (picker == null) { open(); picker = wait(); }
            PersonalActionMemory.Require(picker != null); validate(picker); return picker;
        }
        private static void Reveal(Func<Target> locate, Action require)
        {
            for (int i = 0; i < 24; i++)
            {
                require(); var target = locate(); if (Visible(target)) return;
                PersonalActionMemory.Require(target.Bounds.Width > 1 && target.Bounds.Height > 1);
                var viewport = new Rectangle(Point.Empty, target.Window.Bounds.Size);
                // 목록 밖의 항목은 클릭하지 않고 휠로 표시한 뒤 객체·좌표를 다시 읽습니다.
                var point = new Point(viewport.Width / 2, viewport.Height / 2);
                using (var frame = new Frame { Handle = target.Window.Handle, Bounds = target.Window.Bounds })
                    PersonalActionMemory.Require(Desktop.Apply(frame, new Decision(target.Bounds.Top < 0 || target.Bounds.Left < 0 ? ActionKind.ScrollUp : ActionKind.ScrollDown, "개인 응답 항목 찾기", point) { WheelDelta = 360 }, () => { require(); return false; }));
                Thread.Sleep(100);
            }
            throw new InvalidOperationException("응답 항목을 스크롤 범위 안에서 찾지 못했습니다.");
        }
        internal static void Send(PersonalReplyTicket ticket, NativeMentionInput input, Action require, Action sending, Action reactionSelected)
        {
            var initial = input.Read(); PersonalActionMemory.Require(initial.IsEmpty);
            var room = Desktop.Windows().SingleOrDefault(w => w.Process == ticket.Room.ProcessId && w.Title == ticket.Room.Name && w.Root == w.Handle && w.Owner == IntPtr.Zero);
            PersonalActionMemory.Require(room != null);
            using (var memory = new PersonalActionMemory(ticket.Room.ProcessId))
            {
                Func<ulong> root = () => memory.Room(room.Handle, ticket.Room.ChatId);
                Action contentGuard = () => { require(); root(); PersonalActionMemory.Require(input.Read().IsEmpty); };
                Action guard = contentGuard;
                Action emptyGuard = () => { contentGuard(); PersonalActionMemory.Require(memory.Draft(root()) == null); };
                emptyGuard();
                if (ticket.Action == "reaction")
                {
                    guard = emptyGuard;
                    PersonalActionMemory.Require(PersonalReplyActions.Reactions.ContainsKey(ticket.Choice ?? ""));
                    Func<Target> message = () => { IntPtr window; var bounds = memory.Message(root(), ticket.Room.ChatId, ticket.TargetLogId, ticket.Room.DirectUserId, false, out window); return At(window, bounds, ticket.Room.ProcessId); };
                    Func<Target> opener = () => { IntPtr window; var bounds = memory.Message(root(), ticket.Room.ChatId, ticket.TargetLogId, ticket.Room.DirectUserId, true, out window); return At(window, bounds, ticket.Room.ProcessId); };
                    var popup = OpenOrReusePicker(() => Popup(room, memory, "ReactionPickerWnd"),
                        () => { Reveal(message, guard); Act(message, guard, false); Thread.Sleep(150); Act(opener, guard, true); },
                        () => WaitPopup(room, memory, "ReactionPickerWnd", guard),
                        w => memory.ReactionContext(memory.Dialog(w.Handle, "ReactionPickerWnd"), root(), ticket.Room.ChatId, ticket.TargetLogId));
                    Func<Target> reaction = () => { IntPtr window; var bounds = memory.Reaction(memory.Dialog(popup.Handle, "ReactionPickerWnd"), root(), ticket.Room.ChatId, ticket.TargetLogId, ticket.Choice, out window); return At(window, bounds, ticket.Room.ProcessId); };
                    Reveal(reaction, guard); guard(); sending(); Act(reaction, guard, true);
                    for (int i = 0; i < 20; i++)
                    {
                        if (Popup(room, memory, "ReactionPickerWnd") == null) { reactionSelected(); return; }
                        Thread.Sleep(50);
                    }
                    throw new InvalidOperationException("공감 선택 결과가 불확실합니다. 다시 누르지 않습니다.");
                }
                PersonalActionMemory.Require(ticket.Action == "emoticon" && ticket.Emoticon != null && ticket.Emoticon.Id == ticket.Choice);
                var picker = OpenOrReusePicker(() => Popup(room, memory, "CEmoticonWnd"),
                    () => Act(() => At(room.Handle, memory.Button(root(), "btn_emoticon"), ticket.Room.ProcessId), emptyGuard, true),
                    () => WaitPopup(room, memory, "CEmoticonWnd", emptyGuard),
                    w => memory.EmoticonPicker(memory.Dialog(w.Handle, "CEmoticonWnd"), room.Handle));
                Func<ulong> pickerRoot = () => { ulong value = memory.Dialog(picker.Handle, "CEmoticonWnd"); memory.EmoticonPicker(value, room.Handle); return value; };
                Func<Target> pack = () => { IntPtr window; var bounds = memory.Pack(pickerRoot(), ticket.Emoticon.Pack, out window); return At(window, bounds, ticket.Room.ProcessId); };
                Reveal(pack, emptyGuard); Act(pack, emptyGuard, true); Thread.Sleep(150);
                Func<Target> emot = () => { IntPtr window; var bounds = memory.Emoticon(pickerRoot(), ticket.Emoticon, out window); return At(window, bounds, ticket.Room.ProcessId); };
                ulong owned = 0; bool attempted = false;
                try
                {
                    Reveal(emot, emptyGuard); Act(emot, emptyGuard, true);
                    for (int i = 0; i < 20 && memory.Draft(root()) == null; i++) { require(); Thread.Sleep(50); }
                    PersonalActionMemory.Require(memory.Draft(root(), out owned) == ticket.Choice && owned != 0);
                    Action draftGuard = () => { ulong current; guard(); PersonalActionMemory.Require(memory.Draft(root(), out current) == ticket.Choice && current == owned); };
                    draftGuard(); sending();
                    // 클릭 직전부터는 전송 여부가 불확실할 수 있으므로 초안 삭제·재전송을 하지 않습니다.
                    attempted = true;
                    Act(() => At(room.Handle, memory.Button(root(), "btn_chat_send"), ticket.Room.ProcessId), draftGuard, true);
                    for (int i = 0; i < 20; i++) { Thread.Sleep(50); if (memory.Draft(root()) == null && input.Read().IsEmpty) return; }
                    throw new InvalidOperationException("이모티콘 전송 결과가 불확실합니다. 자동 재전송하지 않습니다.");
                }
                finally
                {
                    if (!attempted && owned != 0)
                    {
                        try
                        {
                            Action ownDraft = () => { ulong current; PersonalActionMemory.Require(input.Read().IsEmpty && memory.Draft(root(), out current) == ticket.Choice && current == owned); };
                            Act(() => { IntPtr window; var bounds = memory.DraftCancel(root(), ticket.Choice, out window); return At(window, bounds, ticket.Room.ProcessId); }, ownDraft, true);
                        }
                        catch (Exception) { /* 사용자가 바꾼 미리보기나 가려진 창은 건드리지 않습니다. */ }
                    }
                }
            }
        }
        [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    }
}
