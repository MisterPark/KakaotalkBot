using System;
using System.Reflection;
using System.Runtime.InteropServices;
using KakaotalkBot;

// 현재 입력 문서와 안내 문구 서식을 읽기만 한다. 입력이나 전송은 하지 않는다.
internal static class MentionPlaceholderProbe
{
    [STAThread]
    private static int Main(string[] args)
    {
        object document = null, range = null, font = null;
        try
        {
            if (args.Length != 1) throw new ArgumentException("방 이름이 필요합니다.");
            var snapshot = KakaoInputMentions.Read(args[0]);
            Console.WriteLine("입력=" + snapshot.Text.Replace("\r", "\\r") + ", IsEmpty=" + snapshot.IsEmpty);
            foreach (uint message in new uint[] { 0xB8, 0xC6 })
            {
                UIntPtr result;
                if (SendMessageTimeout(snapshot.Window, message, UIntPtr.Zero, IntPtr.Zero, 3, 1500, out result) == IntPtr.Zero) throw new Exception("메시지 조회 실패");
                Console.WriteLine("메시지 0x" + message.ToString("X") + "=" + result.ToUInt64());
            }
            Guid iid = new Guid("8CC497C0-A1DF-11CE-8098-00AA0047BE5D");
            Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(snapshot.Window, 0xFFFFFFF0, ref iid, out document));
            range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { 0, Math.Max(0, snapshot.Text.Length - 1) });
            font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
            foreach (string property in new[] { "ForeColor", "BackColor", "Hidden", "Protected", "Name", "Size" })
                Console.WriteLine(property + "=" + font.GetType().InvokeMember(property, BindingFlags.GetProperty, null, font, null));
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { Release(font); Release(range); Release(document); }
    }
    private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW")] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
}
