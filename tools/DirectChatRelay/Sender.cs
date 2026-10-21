using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
namespace DirectChatRelay
{
    sealed class Target
    {
        internal IntPtr Handle; internal int Pid; internal string Title;
        public override string ToString() { return Title + " (PID " + Pid + ")"; }
    }
    static class Sender
    {
        delegate bool EnumProc(IntPtr hwnd, IntPtr data);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr data);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int count);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out int pid);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr FindWindowEx(IntPtr parent,IntPtr after,string cls,string title);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr SendMessageTimeout(IntPtr hwnd,uint msg,IntPtr wp,StringBuilder lp,uint flags,uint timeout,out IntPtr result);
        [DllImport("user32.dll",CharSet=CharSet.Unicode,EntryPoint="SendMessageTimeoutW")] static extern IntPtr SendTimeout(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp,uint flags,uint timeout,out IntPtr result);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
        static string Title(IntPtr hwnd) { var s=new StringBuilder(1024);GetWindowText(hwnd,s,s.Capacity);return s.ToString(); }
        internal static List<Target> Windows()
        {
            var result=new List<Target>();
            EnumWindows((h,p)=>{int pid;GetWindowThreadProcessId(h,out pid);try { if(IsWindowVisible(h)&&Process.GetProcessById(pid).ProcessName=="KakaoTalk"&&FindWindowEx(h,IntPtr.Zero,"RichEdit50W",null)!=IntPtr.Zero)result.Add(new Target{Handle=h,Pid=pid,Title=Title(h)}); }catch(ArgumentException){} return true;},IntPtr.Zero);
            return result;
        }
        static string Read(IntPtr edit)
        {
            var text=new StringBuilder(65536);IntPtr result;
            if(SendMessageTimeout(edit,13,(IntPtr)text.Capacity,text,2,1000,out result)==IntPtr.Zero)throw new InvalidOperationException("입력창 응답 대기");
            return text.ToString();
        }
        // 전용 STA에서만 호출합니다. 확인한 창의 입력 컨트롤에만 붙여넣고 Enter를 보냅니다.
        internal static bool TrySend(Target target,string text,Func<bool> cancelled,ref bool pasted,out string status)
        {
            status="전송 창 대기";int pid;GetWindowThreadProcessId(target.Handle,out pid);
            if(pid!=target.Pid || Title(target.Handle)!=target.Title)return false;
            IntPtr edit=FindWindowEx(target.Handle,IntPtr.Zero,"RichEdit50W",null);
            if(edit==IntPtr.Zero)return false;
            string current=Read(edit);
            if(current.Length!=0 && (!pasted || current!=text)) {status="작성 중인 내용이 있어 대기";return false;}
            SetForegroundWindow(target.Handle);
            if(GetForegroundWindow()!=target.Handle) {status="전송방을 직접 클릭해 주세요";return false;}
            if(cancelled())return false;
            if(current.Length==0)
            {
                Clipboard.SetText(text);IntPtr result;
                if(SendTimeout(edit,0x302,IntPtr.Zero,IntPtr.Zero,2,1000,out result)==IntPtr.Zero)throw new InvalidOperationException("붙여넣기 응답을 확인할 수 없습니다. 입력창을 확인하세요.");
                pasted=true;
            }
            if(Read(edit)!=text) {status="입력 내용이 달라 전송 대기";return false;}
            if(cancelled() || GetForegroundWindow()!=target.Handle)return false;
            if(!PostMessage(edit,0x100,(IntPtr)13,IntPtr.Zero))throw new InvalidOperationException("Enter 입력 실패");
            PostMessage(edit,0x101,(IntPtr)13,IntPtr.Zero);
            pasted=false;status="전송 입력 완료";return true;
        }
    }
}
