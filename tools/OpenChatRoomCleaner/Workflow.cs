using System;
using System.Drawing;
using System.Linq;
namespace OpenChatRoomCleaner
{
    enum RunMode { Observe, Preview, Execute }
    enum Stage { OpenList, SeekTop, Scan, Profile, RemoveMenu, Confirm, ReturnToList, Complete }
    enum ActionKind { Wait, Click, ScrollUp, ScrollDown, Escape, Complete, PreviewComplete, ConfirmRemoval }
    sealed class Decision
    {
        internal ActionKind Kind; internal Point Point; internal string Status;
        internal Decision(ActionKind kind,string status,Point point=default(Point)) { Kind=kind;Status=status;Point=point; }
    }
    // 다음 상태는 실제 입력이 접수된 뒤에만 반영합니다. 인식 실패는 항상 같은 단계에서 대기합니다.
    sealed class Workflow
    {
        readonly RunMode mode;
        internal Workflow(RunMode mode=RunMode.Preview) { this.mode=mode; }
        internal Stage Stage=Stage.OpenList;
        internal int Attempts;
        ulong lastHash; int unchanged; bool scrollIssued;
        internal Decision Decide(Observation o)
        {
            if(Stage==Stage.Complete) return new Decision(ActionKind.Complete,"운영진/본인 외 참여자가 없는 목록 끝 확인");
            if(Stage==Stage.OpenList || Stage==Stage.ReturnToList)
            {
                if(o.ListViewportValid || o.ListValid) { Stage=Stage.SeekTop; ResetScroll(); }
                else if(!o.Members.IsEmpty) return Click(o.Members,"참여자 목록 열기");
                else if(Stage==Stage.ReturnToList && (!o.ProfileMenu.IsEmpty || !o.Remove.IsEmpty)) return new Decision(ActionKind.Escape,"프로필 닫기");
                else return Wait("참여자 목록 또는 채팅방 화면 대기");
            }
            if(Stage==Stage.SeekTop)
            {
                if(!o.ListViewportValid && !o.ListValid) return Wait("스크롤할 목록 영역 인식 대기");
                bool reachedTop=StableAfterScroll(o.ListHash);
                if(reachedTop && o.ListValid) { Stage=Stage.Scan; ResetScroll(); }
                else return new Decision(ActionKind.ScrollUp,
                    o.ListValid?"목록 맨 위 확인 중":"행 판독 전 목록을 맨 위로 이동 중",o.ScrollPoint);
            }
            if(Stage==Stage.Scan)
            {
                if(!o.ListValid) return Wait("참여자 행 인식 재시도");
                var row=o.Rows.FirstOrDefault(r=>!r.Protected);
                if(row!=null) return new Decision(ActionKind.Click,"일반 참여자 프로필 열기",row.Click);
                if(o.Rows.Any(r=>r.Uncertain))return Wait("운영진 배지 여부가 불확실한 행이 있어 재인식 중");
                if(StableAfterScroll(o.ListHash)) { Stage=Stage.Complete; return Decide(o); }
                return new Decision(ActionKind.ScrollDown,"운영진/본인 제외 · 다음 목록 확인",o.ScrollPoint);
            }
            if(Stage==Stage.Profile && !o.ProfileMenu.IsEmpty) return Click(o.ProfileMenu,"내보내기&신고 메뉴 열기");
            if(Stage==Stage.RemoveMenu && !o.Remove.IsEmpty) return Click(o.Remove,"신고 없이 내보내기 선택");
            if(Stage==Stage.Confirm && !o.Confirm.IsEmpty)
            {
                if(mode!=RunMode.Execute)return new Decision(ActionKind.PreviewComplete,"자동 시험 완료: 최종 확인창 도착 · 확인은 누르지 않았습니다");
                var confirm=Click(o.Confirm,"내보내기 확인");confirm.Kind=ActionKind.ConfirmRemoval;return confirm;
            }
            return Wait("예상 화면 대기: "+Stage);
        }
        internal void Applied(Decision d)
        {
            if(d.Kind==ActionKind.ScrollDown || d.Kind==ActionKind.ScrollUp) { scrollIssued=true; return; }
            if(d.Kind!=ActionKind.Click && d.Kind!=ActionKind.ConfirmRemoval) return;
            if(Stage==Stage.Scan) Stage=Stage.Profile;
            else if(Stage==Stage.Profile) Stage=Stage.RemoveMenu;
            else if(Stage==Stage.RemoveMenu) Stage=Stage.Confirm;
            else if(Stage==Stage.Confirm && d.Kind==ActionKind.ConfirmRemoval && mode==RunMode.Execute) { Attempts++; Stage=Stage.ReturnToList; }
        }
        bool StableAfterScroll(ulong hash)
        {
            if(scrollIssued && lastHash==hash) unchanged++; else unchanged=0;
            lastHash=hash; scrollIssued=false; return unchanged>=3;
        }
        void ResetScroll() { unchanged=0;lastHash=0;scrollIssued=false; }
        static Decision Click(Rectangle r,string text) { return new Decision(ActionKind.Click,text,new Point(r.Left+r.Width/2,r.Top+r.Height/2)); }
        static Decision Wait(string text) { return new Decision(ActionKind.Wait,text); }
    }
}
