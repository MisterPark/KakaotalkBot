using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Accessibility;

namespace KakaoMentionInputProbe
{
    // 카카오톡 입력창의 공개 COM 경로만 읽는다. 입력·전송·프로세스 메모리 쓰기는 수행하지 않는다.
    internal static class Program
    {
        private static readonly Guid TextDocument = new Guid("8CC497C0-A1DF-11CE-8098-00AA0047BE5D");
        private static readonly Guid Accessible = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
        private static readonly Guid Dispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static readonly Guid RichEditOle = new Guid("00020D00-0000-0000-C000-000000000046");
        private static readonly StringBuilder Report = new StringBuilder();
        private static bool documentRead;

        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            try
            {
                if ((args.Length != 1 && args.Length != 3) || string.IsNullOrWhiteSpace(args[0]))
                    throw new ArgumentException("사용법: MentionInputProbe.exe \"채팅방 제목\" [user_id 닉네임]");
                long expectedId = 0;
                if (args.Length == 3 && (!long.TryParse(args[1], NumberStyles.None, CultureInfo.InvariantCulture, out expectedId) || expectedId <= 0))
                    throw new ArgumentException("user_id는 양의 64비트 정수여야 합니다.");
                if (args.Length == 3 && (string.IsNullOrWhiteSpace(args[2]) || args[2].Length > 128))
                    throw new ArgumentException("비교할 닉네임은 1~128자로 입력하세요.");
                var rooms = new List<IntPtr>();
                Native.EnumWindows((window, parameter) =>
                {
                    var title = new StringBuilder(512);
                    Native.GetWindowText(window, title, title.Capacity);
                    if (title.ToString() == args[0])
                    {
                        uint pid;
                        Native.GetWindowThreadProcessId(window, out pid);
                        using (var process = Process.GetProcessById((int)pid))
                            if (process.ProcessName.Equals("KakaoTalk", StringComparison.OrdinalIgnoreCase)) rooms.Add(window);
                    }
                    return true;
                }, IntPtr.Zero);
                if (rooms.Count != 1) throw new InvalidOperationException("동일 제목의 카카오톡 창이 정확히 하나 있어야 합니다. 발견: " + rooms.Count);
                uint roomPid;
                Native.GetWindowThreadProcessId(rooms[0], out roomPid);
                using (var process = Process.GetProcessById((int)roomPid))
                {
                    Line("프로세스: " + roomPid + ", 버전: " + process.MainModule.FileVersionInfo.ProductVersion);
                    foreach (ProcessModule module in process.Modules)
                        if (module.ModuleName.Equals("msftedit.dll", StringComparison.OrdinalIgnoreCase))
                            Line("RichEdit 버전: " + module.FileVersionInfo.FileVersion);
                }
                Line("수집 시각: " + DateTimeOffset.Now.ToString("O"));
                Line("방: " + args[0] + ", HWND=" + Hex(rooms[0]));
                var inputs = new List<IntPtr>();
                Native.EnumChildWindows(rooms[0], (window, parameter) =>
                {
                    var name = new StringBuilder(128);
                    Native.GetClassName(window, name, name.Capacity);
                    if (name.ToString().StartsWith("RichEdit", StringComparison.OrdinalIgnoreCase) && Native.GetDlgCtrlID(window) == 1006)
                        inputs.Add(window);
                    return true;
                }, IntPtr.Zero);
                if (inputs.Count != 1) throw new InvalidOperationException("입력 컨트롤을 유일하게 식별하지 못했습니다. 발견: " + inputs.Count);
                Inspect(inputs[0]);
                if (args.Length == 3) Line(MentionMemoryProbe.Read((int)roomPid, expectedId, args[2], inputs[0]));
                string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Diagnostics",
                    DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory, "report.txt");
                File.WriteAllText(path, Report.ToString(), new UTF8Encoding(true));
                Console.Write(Report);
                Console.WriteLine("보고서: " + path);
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(Report.ToString() + "진단 실패: " + error.Message);
                return 1;
            }
        }

        private static void Inspect(IntPtr input)
        {
            Line("입력 HWND=" + Hex(input));
            // WM_GETTEXT는 운영체제가 프로세스 사이 버퍼를 마샬링하는 시스템 메시지다.
            var text = new StringBuilder(4096);
            UIntPtr result;
            if (Native.SendMessageTimeout(input, 0xD, (UIntPtr)text.Capacity, text, 2, 1500, out result) == IntPtr.Zero)
                throw new InvalidOperationException("WM_GETTEXT 시간 초과 또는 실패: " + Marshal.GetLastWin32Error());
            Line("WM_GETTEXT: " + Escape(text.ToString()));
            Line("U+FFFC 개수: " + text.ToString().Count(character => character == '\uFFFC'));
            ProbeWindow(input, 0xFFFFFFF0, Dispatch, "OBJID_NATIVEOM / IDispatch");
            ProbeWindow(input, 0xFFFFFFF0, TextDocument, "OBJID_NATIVEOM / ITextDocument");
            ProbeWindow(input, 0xFFFFFFFC, Accessible, "OBJID_CLIENT / IAccessible");
            ProbeRot(input);
        }

        private static void ProbeWindow(IntPtr window, uint objectId, Guid iid, string label)
        {
            object value = null;
            try
            {
                int hr = Native.AccessibleObjectFromWindow(window, objectId, ref iid, out value);
                Line(label + ": HRESULT=0x" + hr.ToString("X8") + ", 반환=" + (value != null));
                if (hr >= 0 && value != null)
                {
                    ProbeInterfaces(value);
                    var accessible = value as IAccessible;
                    if (accessible != null) InspectAccessible(accessible, 0);
                }
            }
            catch (Exception error) { Line(label + ": " + Error(error)); }
            finally { Release(value); }
        }

        private static void ProbeInterfaces(object value)
        {
            IntPtr unknown = Marshal.GetIUnknownForObject(value);
            try
            {
                foreach (var pair in new[] { new KeyValuePair<string, Guid>("ITextDocument", TextDocument), new KeyValuePair<string, Guid>("IRichEditOle", RichEditOle) })
                {
                    var iid = pair.Value;
                    IntPtr found;
                    int hr = Marshal.QueryInterface(unknown, ref iid, out found);
                    Line("  QueryInterface " + pair.Key + ": 0x" + hr.ToString("X8"));
                    if (found != IntPtr.Zero)
                    {
                        try
                        {
                            if (pair.Key == "ITextDocument" && !documentRead) { documentRead = true; InspectDocument(value); }
                        }
                        finally { Marshal.Release(found); }
                    }
                }
            }
            finally { Marshal.Release(unknown); }
        }

        private static void InspectDocument(object document)
        {
            object range = null, embedded = null;
            try
            {
                range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { 0, 4096 });
                string text = Convert.ToString(range.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, range, null));
                Line("  TOM 입력: " + Escape(text));
                Line("  TOM U+FFFC 개수: " + text.Count(character => character == '\uFFFC'));
                int inspected = 0;
                for (int position = 0; position < text.Length && inspected < 32; position++)
                {
                    if (text[position] != '\uFFFC') continue;
                    Release(range); range = null;
                    range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { position, position + 1 });
                    embedded = range.GetType().InvokeMember("GetEmbeddedObject", BindingFlags.InvokeMethod, null, range, null);
                    Line("  TOM 내부 객체 위치=" + position + ": " + (embedded == null ? "없음" : embedded.GetType().FullName));
                    if (embedded != null) InspectEmbedded(embedded);
                    Release(embedded); embedded = null;
                    inspected++;
                }
            }
            catch (Exception error) { Line("  TOM 조회: " + Error(error)); }
            finally { Release(embedded); Release(range); }
        }

        private static void InspectEmbedded(object embedded)
        {
            IntPtr unknown = Marshal.GetIUnknownForObject(embedded);
            try
            {
                foreach (var pair in new[] {
                    new KeyValuePair<string, string>("IPersist", "0000010C-0000-0000-C000-000000000046"),
                    new KeyValuePair<string, string>("IPersistStorage", "0000010A-0000-0000-C000-000000000046"),
                    new KeyValuePair<string, string>("IPersistStream", "00000109-0000-0000-C000-000000000046"),
                    new KeyValuePair<string, string>("IOleObject", "00000112-0000-0000-C000-000000000046"),
                    new KeyValuePair<string, string>("IDataObject", "0000010E-0000-0000-C000-000000000046"),
                    new KeyValuePair<string, string>("IDispatch", "00020400-0000-0000-C000-000000000046") })
                {
                    Guid iid = new Guid(pair.Value);
                    IntPtr found;
                    int hr = Marshal.QueryInterface(unknown, ref iid, out found);
                    Line("  내부 객체 " + pair.Key + ": 0x" + hr.ToString("X8"));
                    if (found != IntPtr.Zero) Marshal.Release(found);
                }
                var persist = embedded as IPersist;
                if (persist != null)
                {
                    Guid clsid;
                    persist.GetClassID(out clsid);
                    Line("  내부 객체 CLSID: " + clsid);
                }
                var data = embedded as IDataObject;
                if (data != null) InspectFormats(data);
                var ole = embedded as IOleObject;
                if (ole != null)
                {
                    try { Guid clsid; ole.GetUserClassID(out clsid); Line("  IOleObject CLSID: " + clsid); }
                    catch (Exception error) { Line("  IOleObject CLSID: " + Error(error)); }
                    try { string name; ole.GetUserType(1, out name); Line("  IOleObject 사용자 형식: " + Escape(name)); }
                    catch (Exception error) { Line("  IOleObject 사용자 형식: " + Error(error)); }
                    try
                    {
                        IMoniker moniker;
                        ole.GetMoniker(2, 1, out moniker);
                        try { Line("  IOleObject 모니커: " + (moniker != null)); }
                        finally { Release(moniker); }
                    }
                    catch (Exception error) { Line("  IOleObject 모니커: " + Error(error)); }
                }
            }
            finally { Marshal.Release(unknown); }
        }

        private static void InspectFormats(IDataObject data)
        {
            IEnumFORMATETC formats = null;
            try
            {
                formats = data.EnumFormatEtc(DATADIR.DATADIR_GET);
                if (formats == null) { Line("  데이터 형식 열거자 없음"); return; }
                var next = new FORMATETC[1];
                var fetched = new int[1];
                for (int index = 0; index < 32 && formats.Next(1, next, fetched) == 0; index++)
                {
                    FORMATETC format = next[0];
                    try
                    {
                        var name = new StringBuilder(256);
                        Native.GetClipboardFormatName(unchecked((ushort)format.cfFormat), name, name.Capacity);
                        Line("  데이터 형식: " + unchecked((ushort)format.cfFormat) + " " + name + " / " + format.tymed + " / " + format.dwAspect);
                        if ((format.tymed & TYMED.TYMED_HGLOBAL) != 0 && (format.cfFormat == 1 || format.cfFormat == 13))
                        {
                            format.tymed = TYMED.TYMED_HGLOBAL;
                            STGMEDIUM medium;
                            data.GetData(ref format, out medium);
                            try
                            {
                                Line("    GetData 반환 매체: " + medium.tymed);
                                if (medium.tymed != TYMED.TYMED_HGLOBAL) continue;
                                ulong size = Native.GlobalSize(medium.unionmember).ToUInt64();
                                if (size == 0 || size > 64 * 1024) { Line("    원본 크기 제한: " + size); continue; }
                                IntPtr pointer = Native.GlobalLock(medium.unionmember);
                                if (pointer == IntPtr.Zero) throw new IOException("데이터 잠금 실패");
                                try
                                {
                                    var bytes = new byte[(int)size];
                                    Marshal.Copy(pointer, bytes, 0, bytes.Length);
                                    Line("    원본 " + size + "바이트: " + Escape((format.cfFormat == 13 ? Encoding.Unicode : Encoding.Default).GetString(bytes).TrimEnd('\0')));
                                }
                                finally { Native.GlobalUnlock(medium.unionmember); }
                            }
                            finally { Native.ReleaseStgMedium(ref medium); }
                        }
                    }
                    finally { if (format.ptd != IntPtr.Zero) Marshal.FreeCoTaskMem(format.ptd); }
                }
            }
            catch (Exception error) { Line("  데이터 형식 조회: " + Error(error)); }
            finally { Release(formats); }
        }

        private static void InspectAccessible(IAccessible accessible, int depth)
        {
            if (depth > 2) return;
            DescribeAccessible(accessible, 0, depth);
            int count = Math.Min(accessible.accChildCount, 16);
            var children = new object[count];
            int actual;
            if (count == 0 || Native.AccessibleChildren(accessible, 0, count, children, out actual) < 0) return;
            for (int index = 0; index < actual; index++)
            {
                object child = children[index];
                try
                {
                    var childAccessible = child as IAccessible;
                    if (childAccessible != null)
                    {
                        InspectAccessible(childAccessible, depth + 1);
                        ProbeInterfaces(childAccessible);
                    }
                    else if (child is int) DescribeAccessible(accessible, child, depth + 1);
                }
                finally { Release(child); }
            }
        }

        private static void DescribeAccessible(IAccessible accessible, object child, int depth)
        {
            string prefix = new string(' ', depth * 2) + "접근성 child=" + child;
            foreach (var field in new[] { "Name", "Value", "Description", "Role" })
            {
                try
                {
                    object value = field == "Name" ? accessible.get_accName(child) : field == "Value" ? accessible.get_accValue(child) :
                        field == "Description" ? accessible.get_accDescription(child) : accessible.get_accRole(child);
                    Line(prefix + " " + field + "=" + Escape(Convert.ToString(value)));
                }
                catch (Exception error) { Line(prefix + " " + field + ": " + Error(error)); }
            }
        }

        private static void ProbeRot(IntPtr window)
        {
            IRunningObjectTable rot = null;
            IMoniker moniker = null;
            object value = null;
            try
            {
                Marshal.ThrowExceptionForHR(Native.GetRunningObjectTable(0, out rot));
                Marshal.ThrowExceptionForHR(Native.CreateFileMoniker(window.ToInt64().ToString("x"), out moniker));
                int hr = rot.GetObject(moniker, out value);
                Line("ROT HWND 모니커: 0x" + hr.ToString("X8") + ", 반환=" + (value != null));
                if (hr >= 0 && value != null) ProbeInterfaces(value);
            }
            catch (Exception error) { Line("ROT 조회: " + Error(error)); }
            finally { Release(value); Release(moniker); Release(rot); }
        }

        private static string Error(Exception error)
        {
            if (error is TargetInvocationException && error.InnerException != null) error = error.InnerException;
            return error.GetType().Name + " 0x" + error.HResult.ToString("X8") + " " + error.Message;
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        private static string Hex(IntPtr value) { return "0x" + value.ToInt64().ToString("X"); }
        private static string Escape(string value)
        {
            if (value == null) return "<null>";
            return string.Concat(value.Take(512).Select(character => char.IsControl(character) || character == '\uFFFC' ? "\\u" + ((int)character).ToString("X4") : character.ToString()));
        }
        private static void Line(string text) { Report.Append(text).Append("\r\n"); }
    }

    internal static class Native
    {
        internal delegate bool EnumWindowCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")] internal static extern bool EnumWindows(EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] internal static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClipboardFormatName(uint format, StringBuilder name, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wParam, StringBuilder lParam, uint flags, uint timeout, out UIntPtr result);
        [DllImport("oleacc.dll")] internal static extern int AccessibleObjectFromWindow(IntPtr window, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
        [DllImport("oleacc.dll")] internal static extern int AccessibleChildren(IAccessible parent, int start, int count, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] children, out int obtained);
        [DllImport("ole32.dll")] internal static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)] internal static extern int CreateFileMoniker(string path, out IMoniker moniker);
        [DllImport("ole32.dll")] internal static extern void ReleaseStgMedium(ref STGMEDIUM medium);
        [DllImport("kernel32.dll")] internal static extern UIntPtr GlobalSize(IntPtr handle);
        [DllImport("kernel32.dll")] internal static extern IntPtr GlobalLock(IntPtr handle);
        [DllImport("kernel32.dll")] internal static extern bool GlobalUnlock(IntPtr handle);
    }

    [ComImport, Guid("0000010C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPersist
    {
        void GetClassID(out Guid classId);
    }

    [ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleObject
    {
        void SetClientSite(IntPtr site);
        void GetClientSite(out IntPtr site);
        void SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string application, [MarshalAs(UnmanagedType.LPWStr)] string document);
        void Close(uint option);
        void SetMoniker(uint which, IMoniker moniker);
        void GetMoniker(uint assign, uint which, out IMoniker moniker);
        void InitFromData(IDataObject data, [MarshalAs(UnmanagedType.Bool)] bool creation, uint reserved);
        void GetClipboardData(uint reserved, out IDataObject data);
        void DoVerb(int verb, IntPtr message, IntPtr site, int index, IntPtr parent, IntPtr rectangle);
        void EnumVerbs(out IntPtr enumerator);
        void Update();
        void IsUpToDate();
        void GetUserClassID(out Guid classId);
        void GetUserType(uint form, [MarshalAs(UnmanagedType.LPWStr)] out string type);
    }
}
