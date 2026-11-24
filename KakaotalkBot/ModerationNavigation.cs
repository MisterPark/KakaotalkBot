using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OpenChatRoomCleaner;

namespace KakaotalkBot
{
    internal static class ModerationNavigation
    {
        internal static Window Content(Window dialog, IList<Window> windows)
        {
            var panes = windows.Where(w => w.Visible && w.Process == dialog.Process && w.Parent == dialog.Handle &&
                w.ClassName == "EVA_ChildWindow" && w.Bounds.Width >= 200 && w.Bounds.Height >= 120 && dialog.Bounds.Contains(w.Bounds)).ToArray();
            return panes.Length == 1 ? panes[0] : null;
        }
        internal static Window Dialog(Window room, IList<Window> windows)
        {
            var candidates = windows.Where(w => w.Handle != room.Handle && w.Root == w.Handle && w.ClassName == "EVA_Window_Dblclk" &&
                ModerationMenu.Related(w, room, windows, IntPtr.Zero) &&
                (ModerationMenu.Normalize(w.Title) == "오픈채팅관리" || string.IsNullOrWhiteSpace(w.Title))).Where(w =>
                {
                    var pane = Content(w, windows);
                    return pane != null && windows.Any(child => child.Visible && child.Process == w.Process && child.Parent == pane.Handle &&
                        child.ClassName == "_EVA_CustomScrollCtrl" && pane.Bounds.Contains(child.Bounds) && child.Bounds.Height >= 120);
                }).ToArray();
            return candidates.Length == 1 ? candidates[0] : null;
        }
        internal static Point ScrollPoint(Size viewport)
        {
            if (viewport.Width < 200 || viewport.Height < 120) throw new InvalidOperationException("관리 창의 스크롤 영역을 확인하지 못했습니다.");
            // 입력 칸·설정 버튼을 클릭하지 않고, 내용 영역 오른쪽에서 휠만 내립니다.
            return new Point(viewport.Width - 24, viewport.Height / 2);
        }
        internal static bool Search(Action guard, Func<bool> selectVisible, Action scrollDown, Func<bool> waitForChange, int limit = 20)
        {
            for (int step = 0; step <= limit; step++)
            {
                guard();
                if (selectVisible()) return true;
                if (step == limit) return false;
                guard(); scrollDown();
                guard();
                if (!waitForChange())
                {
                    // 끝에 도달했거나 화면이 움직이지 않으면 마지막 화면을 한 번 확인하고 종료합니다.
                    guard(); return selectVisible();
                }
            }
            return false;
        }
    }
}
