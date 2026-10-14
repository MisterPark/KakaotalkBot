using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
namespace OpenChatRoomCleaner
{
    sealed class Window
    {
        internal IntPtr Handle, Parent, Owner, Root; internal uint Process, Thread; internal string Title, ClassName; internal bool Visible; internal Rectangle Bounds;
        public override string ToString() { return Title; }
    }
    sealed class Frame : IDisposable
    {
        internal IntPtr Handle; internal Rectangle Bounds; internal Bitmap Image;
        public void Dispose() { if(Image!=null)Image.Dispose(); }
    }
    static class Desktop
    {
        delegate bool EnumProc(IntPtr h,IntPtr p);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr p);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr p);
        [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h,StringBuilder s,int n);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
        [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h,uint command);
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flag);
        [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
        [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
        [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern void mouse_event(uint flags,uint x,uint y,int data,UIntPtr extra);
        [DllImport("user32.dll")] static extern void keybd_event(byte key,byte scan,uint flags,UIntPtr extra);
        [DllImport("user32.dll")] internal static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] internal static extern bool SetProcessDPIAware();
        [StructLayout(LayoutKind.Sequential)] struct Rect { internal int L,T,R,B; }
        internal static List<Window> Windows(bool includeHidden=false,bool includeChildren=false)
        {
            var windows=new List<Window>();var seen=new HashSet<IntPtr>();var kakao=new HashSet<uint>();
            foreach(var p in Process.GetProcessesByName("KakaoTalk")){using(p)kakao.Add((uint)p.Id);}
            EnumProc visit=(h,p)=>
            {
                uint pid;uint thread=GetWindowThreadProcessId(h,out pid);bool visible=IsWindowVisible(h)&&!IsIconic(h);
                if(kakao.Contains(pid)&&seen.Add(h)&&(includeHidden||visible))
                {
                    var title=new StringBuilder(256);var name=new StringBuilder(128);Rect rect;
                    GetWindowText(h,title,title.Capacity);GetClassName(h,name,name.Capacity);GetWindowRect(h,out rect);
                    windows.Add(new Window{Handle=h,Process=pid,Thread=thread,Title=title.ToString(),ClassName=name.ToString(),
                        Parent=GetParent(h),Owner=GetWindow(h,4),Root=GetAncestor(h,2),Visible=visible,Bounds=new Rectangle(rect.L,rect.T,rect.R-rect.L,rect.B-rect.T)});
                }
                return true;
            };
            EnumWindows((h,p)=>{uint pid;GetWindowThreadProcessId(h,out pid);if(kakao.Contains(pid)){visit(h,p);if(includeChildren)EnumChildWindows(h,visit,p);}return true;},IntPtr.Zero);
            return windows;
        }
        internal static bool IsCaptureWindow(Window window)
        {
            // GetParent는 소유 팝업에서도 0이 아닙니다. 실제 자식 여부는 GA_ROOT로 판별합니다.
            return window.Visible && window.Root==window.Handle && window.Handle!=IntPtr.Zero;
        }
        internal static bool IsParticipantPopup(Window window,Window room,List<Window> windows)
        {
            if(!IsCaptureWindow(window)||window.Handle==room.Handle||window.Process!=room.Process
                ||window.ClassName!="EVA_Window_Dblclk"||window.Bounds.Width<265||window.Bounds.Width>285)return false;
            var visited=new HashSet<IntPtr>();IntPtr owner=window.Owner;
            while(owner!=IntPtr.Zero&&visited.Add(owner))
            {
                if(owner==room.Handle)return true;
                var parent=windows.Find(w=>w.Handle==owner && w.Process==room.Process);
                if(parent==null)return false;owner=parent.Owner;
            }
            return false;
        }
        // 실제 관측: 프로필(300x600) → 참여자 목록(270 폭) → 선택 채팅방.
        // 클래스 하나만으로는 채팅방과 구분되지 않으므로 크기 비율과 소유 관계도 검사합니다.
        internal static bool IsProfileWindow(Window w,Window room,List<Window> windows)
        {
            if(!w.Visible || w.Process!=room.Process || w.ClassName!="EVA_Window_Dblclk")return false;
            double scale=w.Bounds.Width/300.0;
            if(scale<0.75 || scale>2.5 || Math.Abs(w.Bounds.Height-600*scale)>6*scale)return false;
            var owner=windows.Find(x=>x.Handle==w.Owner && x.Visible && x.Process==room.Process);
            if(owner==null || owner.ClassName!="EVA_Window_Dblclk" || Math.Abs(owner.Bounds.Width-270*scale)>6*scale)return false;
            IntPtr current=owner.Handle;var visited=new HashSet<IntPtr>();
            while(current!=IntPtr.Zero && visited.Add(current))
            {
                if(current==room.Handle)return true;
                var ancestor=windows.Find(x=>x.Handle==current);if(ancestor==null)return false;
                current=ancestor.Owner;
            }
            return false;
        }
        internal static bool OwnedBy(IntPtr h,IntPtr room)
        {
            for(int i=0;i<16 && h!=IntPtr.Zero;i++,h=GetWindow(h,4)) if(h==room)return true;
            return false;
        }
        internal static bool IsForegroundProcess(uint pid) { uint actual;GetWindowThreadProcessId(GetForegroundWindow(),out actual);return actual==pid; }
        internal static Frame Capture(IntPtr h)
        {
            Rect r; if(!IsWindowVisible(h)||IsIconic(h)||!GetWindowRect(h,out r))return null;
            var b=new Rectangle(r.L,r.T,r.R-r.L,r.B-r.T);
            if(b.Width<10||b.Height<10||b.Width>2500||b.Height>1800)return null;
            var image=new Bitmap(b.Width,b.Height);
            try { using(var g=Graphics.FromImage(image))g.CopyFromScreen(b.Location,Point.Empty,b.Size); }
            catch {image.Dispose();throw;}
            return new Frame{Handle=h,Bounds=b,Image=image};
        }
        internal static bool Apply(Frame f,Decision d,Func<bool> cancelled,bool allowRemoval=false)
        {
            if(cancelled() || (d.Kind==ActionKind.ConfirmRemoval && !allowRemoval))return false;
            Rect current; if(!GetWindowRect(f.Handle,out current) || new Rectangle(current.L,current.T,current.R-current.L,current.B-current.T)!=f.Bounds)return false;
            var p=new Point(f.Bounds.Left+d.Point.X,f.Bounds.Top+d.Point.Y);
            if(d.Kind==ActionKind.Escape)
            {
                if(GetAncestor(GetForegroundWindow(),2)!=GetAncestor(f.Handle,2))return false;
                if(cancelled())return false;keybd_event(27,0,0,UIntPtr.Zero);keybd_event(27,0,2,UIntPtr.Zero);return true;
            }
            if(!f.Bounds.Contains(p)||GetAncestor(WindowFromPoint(p),2)!=GetAncestor(f.Handle,2))return false;
            if(cancelled()||!SetCursorPos(p.X,p.Y))return false;
            if(d.Kind==ActionKind.Click || d.Kind==ActionKind.ConfirmRemoval) {mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);}
            else if(d.Kind==ActionKind.ScrollUp || d.Kind==ActionKind.ScrollDown)mouse_event(0x0800,0,0,d.Kind==ActionKind.ScrollUp?360:-360,UIntPtr.Zero);
            else return false;
            return true;
        }
    }
}
