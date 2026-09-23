using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace KakaotalkBot
{
    public sealed class InputMention
    {
        public int Position { get; internal set; }
        public long UserId { get; internal set; }
        public string DisplayText { get; internal set; }
        internal ulong Address, Site;
    }

    public sealed class InputMentionSnapshot
    {
        public string RoomTitle { get; internal set; }
        public string Text { get; internal set; }
        public IReadOnlyList<InputMention> Mentions { get; internal set; }
        public bool IsEmpty { get { return Mentions.Count == 0 && (Text == "\r" || IsPlaceholder); } }
        internal bool IsPlaceholder;
        internal IntPtr Window;
        internal int ProcessId;
        internal long ProcessStart;
    }

    /// <summary>현재 입력 중인 멘션을 조회한다. 검증된 카카오톡·RichEdit 버전에서만 작동한다.</summary>
    public static class KakaoInputMentions
    {
        private const string KakaoVersion = "26.8.1.5315";
        private const string RichEditVersion = "10.0.19041.4522";
        private const ulong MentionVtableRva = 0x292FB50;
        private static readonly Guid MentionClass = new Guid("459e8c62-070f-42b0-b92b-64dd59af0ebe");
        private static readonly object Gate = new object();
        private static InputMentionSnapshot cached;

        internal static InputMentionSnapshot EditOwnedText(InputMentionSnapshot expected, string append, bool clear)
        {
            lock (Gate)
            using (var context = Open(expected.RoomTitle, false))
            {
                var current = Read(context);
                if (!SameSnapshot(current, expected)) throw new InvalidOperationException("입력 내용이 바뀌어 편집을 중단했습니다.");
                object document = null, range = null;
                try
                {
                    Guid iid = new Guid("8CC497C0-A1DF-11CE-8098-00AA0047BE5D");
                    Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(context.Input, 0xFFFFFFF0, ref iid, out document));
                    int end = Math.Max(0, current.Text.Length - 1);
                    range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { clear ? 0 : end, end });
                    range.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, range, new object[] { clear ? "" : append });
                }
                finally { Release(range); Release(document); }
                var result = Read(context);
                string text = clear ? "\r" : current.Text.Substring(0, Math.Max(0, current.Text.Length - 1)) + append + "\r";
                if (result.Text != text) throw new InvalidOperationException("입력 편집 결과가 예상과 다릅니다.");
                return result;
            }
        }

        internal static bool SameSnapshot(InputMentionSnapshot left, InputMentionSnapshot right)
        {
            return left != null && right != null && left.Window == right.Window && left.ProcessId == right.ProcessId &&
                left.ProcessStart == right.ProcessStart && left.Text == right.Text && left.IsPlaceholder == right.IsPlaceholder && left.Mentions.Count == right.Mentions.Count &&
                left.Mentions.Zip(right.Mentions, (a, b) => a.Address == b.Address && a.Site == b.Site && a.Position == b.Position &&
                    a.UserId == b.UserId && a.DisplayText == b.DisplayText).All(value => value);
        }

        public static InputMentionSnapshot Read(string roomTitle)
        {
            lock (Gate)
            {
                using (var context = Open(roomTitle, false)) return Read(context);
            }
        }

        /// <summary>기존 멘션 하나의 ID만 변경한다. 표시 이름은 유지하며 전송은 호출자가 별도로 수행한다.</summary>
        public static InputMentionSnapshot ReplaceUserId(string roomTitle, int position, long expectedUserId, long replacementUserId,
            string expectedInputText, string expectedDisplayText)
        {
            if (expectedUserId <= 0 || replacementUserId <= 0 || expectedUserId == replacementUserId || expectedInputText == null || expectedDisplayText == null)
                throw new ArgumentException("현재 입력 내용·표시 이름과 서로 다른 양의 사용자 ID가 필요합니다.");
            lock (Gate)
            {
                using (var context = Open(roomTitle, true))
                {
                    var before = Read(context);
                    var mention = before.Mentions.SingleOrDefault(item => item.Position == position);
                    if (before.Text != expectedInputText || before.Mentions.Count != 1 || mention == null || mention.UserId != expectedUserId || mention.DisplayText != expectedDisplayText)
                        throw new InvalidOperationException("현재 입력 내용 또는 대상이 예상과 다릅니다. ID를 변경하지 않았습니다.");
                    RequireSame(context, before, mention, expectedUserId);
                    try
                    {
                        WriteId(context, mention.Address, replacementUserId);
                        RequireSame(context, before, mention, replacementUserId);
                        return Read(context);
                    }
                    catch
                    {
                        var current = ReadObject(context, mention.Address);
                        if (current == null || current.Site != mention.Site || current.Position != position || current.DisplayText != expectedDisplayText ||
                            (current.UserId != expectedUserId && current.UserId != replacementUserId))
                            throw new InvalidOperationException("변경 검증과 원상 복원에 실패했습니다. 전송하지 말고 입력 멘션을 다시 선택해 주세요.");
                        if (current.UserId == replacementUserId) WriteId(context, mention.Address, expectedUserId);
                        RequireSame(context, before, mention, expectedUserId);
                        throw;
                    }
                }
            }
        }

        /// <summary>한 ID 필드의 쓰기·읽기·원상 복원을 검사한다. 표시 이름과 입력 내용은 바꾸거나 전송하지 않는다.</summary>
        public static string VerifyUserIdWriteAndRestore(string roomTitle, int position, long expectedUserId, long replacementUserId)
        {
            if (expectedUserId <= 0 || replacementUserId <= 0 || expectedUserId == replacementUserId)
                throw new ArgumentException("서로 다른 양의 사용자 ID가 필요합니다.");
            lock (Gate)
            {
                using (var context = Open(roomTitle, true))
                {
                    var before = Read(context);
                    var mention = before.Mentions.SingleOrDefault(item => item.Position == position);
                    if (mention == null || mention.UserId != expectedUserId) throw new InvalidOperationException("선택 위치의 현재 사용자 ID가 예상과 다릅니다.");
                    bool attempted = false;
                    try
                    {
                        RequireSame(context, before, mention, expectedUserId);
                        attempted = true;
                        WriteId(context, mention.Address, replacementUserId);
                        RequireSame(context, before, mention, replacementUserId);
                    }
                    finally
                    {
                        if (attempted)
                        {
                            // 객체가 교체되면 이전 주소에 쓰지 않는다. 복원 실패를 성공으로 숨기지 않는다.
                            var current = ReadObject(context, mention.Address);
                            if (current == null || current.Site != mention.Site || current.Position != mention.Position || current.DisplayText != mention.DisplayText)
                                throw new InvalidOperationException("복원 실패: 입력 객체가 변경되었습니다. 전송하지 말고 해당 입력 멘션을 지워 주세요.");
                            if (current.UserId != expectedUserId)
                            {
                                if (current.UserId != replacementUserId) throw new InvalidOperationException("복원 실패: ID가 외부에서 변경되었습니다. 해당 입력 멘션을 다시 선택해 주세요.");
                                WriteId(context, mention.Address, expectedUserId);
                            }
                            RequireSame(context, before, mention, expectedUserId);
                        }
                    }
                    return "ID 쓰기·읽기 확인 및 원래 ID 복원 완료. 송신 결과는 아직 검증하지 않았습니다.";
                }
            }
        }

        private static void RequireSame(Context context, InputMentionSnapshot before, InputMention original, long expectedId)
        {
            var now = Read(context);
            var current = now.Mentions.SingleOrDefault(item => item.Position == original.Position);
            if (now.Text != before.Text || now.Mentions.Count != before.Mentions.Count || current == null ||
                current.Address != original.Address || current.Site != original.Site || current.DisplayText != original.DisplayText || current.UserId != expectedId)
                throw new InvalidOperationException("검사 도중 입력 내용 또는 멘션 객체가 바뀌었습니다.");
        }

        private static void WriteId(Context context, ulong address, long value)
        {
            byte[] bytes = BitConverter.GetBytes(value);
            UIntPtr written;
            if (!WriteProcessMemory(context.Handle, (IntPtr)(long)(address + 0xC0), bytes, (UIntPtr)bytes.Length, out written) || written.ToUInt64() != 8)
                throw new IOException("사용자 ID 쓰기 실패: " + Marshal.GetLastWin32Error());
        }

        private static InputMentionSnapshot Read(Context context)
        {
            var document = ReadDocument(context.Input);
            var objects = new List<InputMention>();
            if (document.Positions.Count > 0)
            {
                if (cached != null && cached.Window == context.Input && cached.ProcessId == context.Pid && cached.ProcessStart == context.Start)
                    foreach (var previous in cached.Mentions)
                    {
                        var current = ReadObject(context, previous.Address);
                        if (current != null && document.Positions.Contains(current.Position)) objects.Add(current);
                    }
                if (!PositionsMatch(document.Positions, objects))
                {
                    objects.Clear();
                    ScanObjects(context, document.Positions, objects);
                }
                if (!PositionsMatch(document.Positions, objects)) throw new InvalidOperationException("현재 TOM 객체와 메모리 객체를 일대일로 확인하지 못했습니다.");
            }
            var after = ReadDocument(context.Input);
            if (after.Text != document.Text || after.IsPlaceholder != document.IsPlaceholder || !after.Positions.SequenceEqual(document.Positions)) throw new InvalidOperationException("읽는 동안 입력 내용이 바뀌었습니다. 다시 조회하세요.");
            foreach (var item in objects)
            {
                var check = ReadObject(context, item.Address);
                if (check == null || check.Site != item.Site || check.Position != item.Position || check.UserId != item.UserId || check.DisplayText != item.DisplayText)
                    throw new InvalidOperationException("읽는 동안 멘션 객체가 바뀌었습니다. 다시 조회하세요.");
            }
            cached = new InputMentionSnapshot { RoomTitle = context.Title, Text = document.Text, Window = context.Input,
                ProcessId = context.Pid, ProcessStart = context.Start, IsPlaceholder = document.IsPlaceholder,
                Mentions = new ReadOnlyCollection<InputMention>(objects.OrderBy(item => item.Position).ToList()) };
            return cached;
        }

        internal static bool IsEmptyPlaceholder(string text, int foreColor, bool canUndo)
        {
            // 지원 버전의 안내 문구는 #949494이며 실행 취소 이력이 없다.
            // EM_GETMODIFY는 안내 문구만 있어도 참이므로 빈 입력 판정에 사용하지 않는다.
            // 직접 작성한 같은 문구나 서식이 섞인 입력은 초안으로 보존한다.
            return text == "메시지 입력\r" && foreColor == 0x949494 && !canUndo;
        }

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);

        private static bool PositionsMatch(List<int> positions, List<InputMention> mentions)
        {
            return positions.OrderBy(value => value).SequenceEqual(mentions.Select(item => item.Position).OrderBy(value => value));
        }

        private static InputMention ReadObject(Context context, ulong address)
        {
            byte[] bytes = Bytes(context, address, 0xD8);
            if (bytes == null || BitConverter.ToUInt64(bytes, 0) != context.Vtable) return null;
            ulong site = BitConverter.ToUInt64(bytes, 0x80);
            byte[] owner = Bytes(context, site, 0x48);
            if (owner == null || BitConverter.ToUInt64(owner, 0x30) != context.Editor || BitConverter.ToUInt64(owner, 0x38) != address) return null;
            ulong siteVtable = BitConverter.ToUInt64(owner, 0);
            if (siteVtable < context.RichEditBase || siteVtable - context.RichEditBase >= context.RichEditSize) return null;
            int position = BitConverter.ToInt32(owner, 0x40);
            long userId = BitConverter.ToInt64(bytes, 0xC0);
            if (position < 0 || userId <= 0) return null;
            ulong length = BitConverter.ToUInt64(bytes, 0x48), capacity = BitConverter.ToUInt64(bytes, 0x50);
            if (length == 0 || length > 256 || capacity < length || capacity > 4096) return null;
            byte[] label = capacity <= 7 ? bytes.Skip(0x38).Take((int)length * 2).ToArray() : Bytes(context, BitConverter.ToUInt64(bytes, 0x38), (int)length * 2);
            if (label == null) return null;
            string display = Encoding.Unicode.GetString(label);
            if (!display.StartsWith("@", StringComparison.Ordinal) || display.IndexOf('\0') >= 0) return null;
            return new InputMention { Address = address, Site = site, Position = position, UserId = userId, DisplayText = display };
        }

        private static void ScanObjects(Context context, List<int> positions, List<InputMention> objects)
        {
            byte[] pattern = BitConverter.GetBytes(context.Vtable);
            ulong address = 0, scanned = 0;
            var clock = Stopwatch.StartNew();
            var found = new HashSet<ulong>();
            while (address < 0x7FFFFFFFFFFF)
            {
                MemoryInformation info;
                if (VirtualQueryEx(context.Handle, (IntPtr)(long)address, out info, (UIntPtr)Marshal.SizeOf(typeof(MemoryInformation))) == UIntPtr.Zero) break;
                ulong start = (ulong)info.BaseAddress.ToInt64(), size = info.RegionSize.ToUInt64();
                if (info.State == 0x1000 && info.Type == 0x20000 && (info.Protect & 0x101) == 0 && (info.Protect & 0xCC) != 0)
                {
                    for (ulong offset = 0; offset < size; offset += 1024 * 1024)
                    {
                        if (scanned >= 1024UL * 1024 * 1024 || clock.ElapsedMilliseconds > 20000) throw new IOException("멘션 객체 검색 한도를 초과했습니다.");
                        int length = (int)Math.Min(1024UL * 1024 + 7, size - offset);
                        byte[] bytes = Bytes(context, start + offset, length);
                        scanned += (ulong)length;
                        if (bytes == null) continue;
                        int index = 0;
                        while (index <= bytes.Length - 8)
                        {
                            index = Array.IndexOf(bytes, pattern[0], index, bytes.Length - 7 - index);
                            if (index < 0) break;
                            ulong candidate = start + offset + (ulong)index;
                            if (candidate % 8 == 0 && BitConverter.ToUInt64(bytes, index) == context.Vtable && found.Add(candidate))
                            {
                                var mention = ReadObject(context, candidate);
                                if (mention != null && positions.Contains(mention.Position)) objects.Add(mention);
                            }
                            index++;
                        }
                    }
                }
                if (size == 0 || start + size <= address) break;
                address = start + size;
            }
        }

        private static Context Open(string title, bool writable)
        {
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("채팅방 제목이 필요합니다.");
            if (!Environment.Is64BitProcess || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("64비트 STA 스레드에서 호출하세요.");
            var rooms = new List<IntPtr>();
            EnumWindows((window, parameter) =>
            {
                var text = new StringBuilder(512);
                GetWindowText(window, text, text.Capacity);
                if (text.ToString() == title)
                {
                    uint pid; GetWindowThreadProcessId(window, out pid);
                    using (var process = Process.GetProcessById((int)pid))
                        if (process.ProcessName.Equals("KakaoTalk", StringComparison.OrdinalIgnoreCase)) rooms.Add(window);
                }
                return true;
            }, IntPtr.Zero);
            if (rooms.Count != 1) throw new InvalidOperationException("제목이 일치하는 카카오톡 창이 정확히 하나 있어야 합니다.");
            var inputs = new List<IntPtr>();
            EnumChildWindows(rooms[0], (window, parameter) =>
            {
                var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                if (name.ToString().Equals("RichEdit50W", StringComparison.OrdinalIgnoreCase) && GetDlgCtrlID(window) == 1006) inputs.Add(window);
                return true;
            }, IntPtr.Zero);
            if (inputs.Count != 1) throw new InvalidOperationException("입력창을 유일하게 확인하지 못했습니다.");
            uint ownerPid; GetWindowThreadProcessId(inputs[0], out ownerPid);
            var context = new Context { Input = inputs[0], Pid = (int)ownerPid, Title = title };
            try
            {
                using (var process = Process.GetProcessById(context.Pid))
                {
                    context.Start = process.StartTime.ToUniversalTime().Ticks;
                    if (process.MainModule.FileVersionInfo.ProductVersion != KakaoVersion) throw new NotSupportedException("검증되지 않은 카카오톡 버전입니다.");
                    context.Vtable = (ulong)process.MainModule.BaseAddress.ToInt64() + MentionVtableRva;
                    var richEdit = process.Modules.Cast<ProcessModule>().Single(module => module.ModuleName.Equals("msftedit.dll", StringComparison.OrdinalIgnoreCase));
                    if (richEdit.FileVersionInfo.FileVersion.Split(' ')[0] != RichEditVersion) throw new NotSupportedException("검증되지 않은 RichEdit 버전입니다.");
                    context.RichEditBase = (ulong)richEdit.BaseAddress.ToInt64(); context.RichEditSize = (ulong)richEdit.ModuleMemorySize;
                }
                context.Handle = OpenProcess(writable ? 0x438u : 0x410u, false, context.Pid);
                if (context.Handle == IntPtr.Zero) throw new IOException("카카오톡 메모리 열기 실패: " + Marshal.GetLastWin32Error());
                bool wow64;
                if (!IsWow64Process(context.Handle, out wow64) || wow64) throw new NotSupportedException("64비트 카카오톡만 지원합니다.");
                ulong root = (ulong)GetWindowLongPtr(context.Input, 0).ToInt64();
                context.Editor = U64(context, root + 0x48);
                ulong getter = U64(context, context.Vtable + 15 * 8);
                byte[] code = Bytes(context, getter, 13);
                if (code == null || code[0] != 0x0F || code[1] != 0x10 || code[2] != 0x05 || code[7] != 0x33 || code[8] != 0xC0 || code[9] != 0x0F || code[10] != 0x11 || code[11] != 0x02 || code[12] != 0xC3)
                    throw new NotSupportedException("멘션 객체의 클래스 검사 함수가 예상과 다릅니다.");
                byte[] guid = Bytes(context, (ulong)((long)getter + 7 + BitConverter.ToInt32(code, 3)), 16);
                if (guid == null || new Guid(guid) != MentionClass) throw new NotSupportedException("멘션 객체의 클래스가 일치하지 않습니다.");
                return context;
            }
            catch { context.Dispose(); throw; }
        }

        private static Document ReadDocument(IntPtr window)
        {
            object document = null, range = null, embedded = null, font = null;
            try
            {
                Guid iid = new Guid("8CC497C0-A1DF-11CE-8098-00AA0047BE5D");
                Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(window, 0xFFFFFFF0, ref iid, out document));
                range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { 0, 4097 });
                string text = Convert.ToString(range.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, range, null));
                if (text.Length > 4096) throw new NotSupportedException("4,096자를 넘는 입력은 검사하지 않습니다.");
                var result = new Document { Text = text };
                if (text == "메시지 입력\r")
                {
                    // 마지막 문단 기호를 제외한 안내 문구 전체의 글자색을 조회한다.
                    Release(range); range = null;
                    range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { 0, text.Length - 1 });
                    font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
                    int foreColor = Convert.ToInt32(font.GetType().InvokeMember("ForeColor", BindingFlags.GetProperty, null, font, null));
                    UIntPtr canUndo;
                    if (SendMessageTimeout(window, 0xC6, UIntPtr.Zero, IntPtr.Zero, 3, 1500, out canUndo) == IntPtr.Zero)
                        throw new InvalidOperationException("입력창 안내 문구 상태를 확인하지 못했습니다.");
                    result.IsPlaceholder = IsEmptyPlaceholder(text, foreColor, canUndo != UIntPtr.Zero);
                }
                for (int position = 0; position < text.Length; position++)
                {
                    if (text[position] != '\uFFFC') continue;
                    if (result.Positions.Count >= 32) throw new NotSupportedException("32개를 넘는 객체는 검사하지 않습니다.");
                    Release(range); range = null;
                    range = document.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, document, new object[] { position, position + 1 });
                    embedded = range.GetType().InvokeMember("GetEmbeddedObject", BindingFlags.InvokeMethod, null, range, null);
                    var ole = embedded as IOleObject;
                    Guid classId;
                    if (ole == null) throw new NotSupportedException("지원하지 않는 입력 객체입니다.");
                    ole.GetUserClassID(out classId);
                    if (classId != MentionClass) throw new NotSupportedException("멘션 이외의 입력 객체가 포함되어 있습니다.");
                    result.Positions.Add(position);
                    Release(embedded); embedded = null;
                }
                return result;
            }
            finally { Release(embedded); Release(font); Release(range); Release(document); }
        }

        private static byte[] Bytes(Context context, ulong address, int length)
        {
            if (address < 0x10000 || address > 0x7FFFFFFFFFFF || length <= 0) return null;
            var bytes = new byte[length]; UIntPtr read;
            return ReadProcessMemory(context.Handle, (IntPtr)(long)address, bytes, (UIntPtr)length, out read) && read.ToUInt64() == (ulong)length ? bytes : null;
        }
        private static ulong U64(Context context, ulong address)
        {
            byte[] bytes = Bytes(context, address, 8);
            if (bytes == null) throw new IOException("입력 객체 포인터를 읽지 못했습니다.");
            return BitConverter.ToUInt64(bytes, 0);
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        private sealed class Document { internal string Text; internal bool IsPlaceholder; internal readonly List<int> Positions = new List<int>(); }
        private sealed class Context : IDisposable
        {
            internal string Title; internal int Pid; internal long Start; internal IntPtr Input, Handle;
            internal ulong Vtable, Editor, RichEditBase, RichEditSize;
            public void Dispose() { if (Handle != IntPtr.Zero) { CloseHandle(Handle); Handle = IntPtr.Zero; } }
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct MemoryInformation
        {
            internal IntPtr BaseAddress, AllocationBase; internal uint AllocationProtect; internal ushort PartitionId;
            internal UIntPtr RegionSize; internal uint State, Protect, Type;
        }
        private delegate bool EnumCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool IsWow64Process(IntPtr process, out bool wow64);
        [DllImport("kernel32.dll")] private static extern UIntPtr VirtualQueryEx(IntPtr process, IntPtr address, out MemoryInformation info, UIntPtr length);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, [Out] byte[] bytes, UIntPtr length, out UIntPtr read);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool WriteProcessMemory(IntPtr process, IntPtr address, byte[] bytes, UIntPtr length, out UIntPtr written);

        // 앞선 메서드는 호출하지 않지만 GetUserClassID의 COM 슬롯 순서를 보존해야 한다.
        [ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleObject
        {
            void SetClientSite(IntPtr site); void GetClientSite(out IntPtr site);
            void SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string application, [MarshalAs(UnmanagedType.LPWStr)] string document);
            void Close(uint option); void SetMoniker(uint which, IntPtr moniker); void GetMoniker(uint assign, uint which, out IntPtr moniker);
            void InitFromData(IntPtr data, [MarshalAs(UnmanagedType.Bool)] bool creation, uint reserved); void GetClipboardData(uint reserved, out IntPtr data);
            void DoVerb(int verb, IntPtr message, IntPtr site, int index, IntPtr parent, IntPtr rectangle);
            void EnumVerbs(out IntPtr enumerator); void Update(); void IsUpToDate(); void GetUserClassID(out Guid classId);
        }
    }
}
