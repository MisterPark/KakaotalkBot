using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KakaoMentionInputProbe
{
    internal static class MentionMemoryProbe
    {
        // 알려진 ID의 위치와 인접 포인터만 조사한다. 원본 덤프와 쓰기 권한은 사용하지 않는다.
        internal static string Read(int pid, long userId, string nickname, IntPtr input)
        {
            if (!Environment.Is64BitProcess) throw new PlatformNotSupportedException("64비트 진단 도구가 필요합니다.");
            var report = new StringBuilder();
            IntPtr process = OpenProcess(0x410, false, pid);
            if (process == IntPtr.Zero) throw new IOException("읽기용 프로세스 열기 실패: " + Marshal.GetLastWin32Error());
            try
            {
                bool wow64;
                if (!IsWow64Process(process, out wow64) || wow64) throw new PlatformNotSupportedException("64비트 카카오톡만 분석할 수 있습니다.");
                var modules = new List<ModuleRange>();
                using (var target = Process.GetProcessById(pid))
                {
                    if (!target.ProcessName.Equals("KakaoTalk", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("대상 프로세스 불일치");
                    foreach (ProcessModule module in target.Modules)
                        modules.Add(new ModuleRange { Start = (ulong)module.BaseAddress.ToInt64(), Size = (ulong)module.ModuleMemorySize, Name = module.ModuleName });
                }
                var idBytes = BitConverter.GetBytes(userId);
                var nameBytes = Encoding.Unicode.GetBytes(nickname);
                var ids = new List<ulong>();
                var names = new List<ulong>();
                var objects = new List<ulong>();
                ulong address = 0, scanned = 0;
                int failedReads = 0;
                var clock = Stopwatch.StartNew();
                while (address < 0x7FFFFFFFFFFF && scanned < 1024UL * 1024 * 1024 && clock.ElapsedMilliseconds < 20000 && ids.Count < 512 && names.Count < 512)
                {
                    MemoryInformation info;
                    if (VirtualQueryEx(process, (IntPtr)(long)address, out info, (UIntPtr)Marshal.SizeOf(typeof(MemoryInformation))) == UIntPtr.Zero) break;
                    ulong start = (ulong)info.BaseAddress.ToInt64(), size = info.RegionSize.ToUInt64();
                    // 쓰기 가능한 커밋된 전용 메모리만 검색하여 이미지·파일 매핑은 제외한다.
                    if (info.State == 0x1000 && info.Type == 0x20000 && (info.Protect & 0x101) == 0 && (info.Protect & 0xCC) != 0)
                    {
                        for (ulong offset = 0; offset < size && scanned < 1024UL * 1024 * 1024 && clock.ElapsedMilliseconds < 20000; offset += 1024 * 1024)
                        {
                            int length = (int)Math.Min(1024UL * 1024 + (ulong)Math.Max(idBytes.Length, nameBytes.Length) - 1, size - offset);
                            byte[] bytes = ReadBytes(process, start + offset, length);
                            if (bytes == null) { failedReads++; continue; }
                            scanned += (ulong)length;
                            Find(bytes, idBytes, start + offset, ids);
                            if (nameBytes.Length > 0) Find(bytes, nameBytes, start + offset, names);
                        }
                    }
                    if (size == 0 || start + size <= address) break;
                    address = start + size;
                }
                report.AppendFormat("메모리 비교 ID={0}, 이름={1}\r\n검색={2} MiB, 시간={3} ms, 실패 청크={4}\r\n", userId, nickname, scanned / 1024 / 1024, clock.ElapsedMilliseconds, failedReads);
                report.AppendFormat("ID 후보={0}, 이름 후보={1}, 제한 도달={2}\r\n", ids.Count, names.Count, scanned >= 1024UL * 1024 * 1024 || clock.ElapsedMilliseconds >= 20000 || ids.Count >= 512 || names.Count >= 512);
                foreach (ulong hit in ids)
                {
                    report.AppendFormat("ID 0x{0:X}", hit);
                    ulong begin = hit > 256 ? (hit - 256) & ~7UL : 0;
                    byte[] nearby = ReadBytes(process, begin, 768);
                    if (nearby != null)
                    {
                        for (int offset = 0; offset + 8 <= nearby.Length; offset += 8)
                        {
                            ulong pointer = BitConverter.ToUInt64(nearby, offset);
                            var module = modules.FirstOrDefault(item => pointer >= item.Start && pointer - item.Start < item.Size);
                            if (module != null && module.Name.Equals("KakaoTalk.exe", StringComparison.OrdinalIgnoreCase))
                            {
                                report.AppendFormat("; [{0:+#;-#;0}]=KakaoTalk.exe+0x{1:X}", (long)(begin + (ulong)offset) - (long)hit, pointer - module.Start);
                                if (begin + (ulong)offset < hit && names.Any(value => value >= begin && value < begin + 512))
                                {
                                    byte[] table = ReadBytes(process, pointer, 16 * 8);
                                    if (table != null)
                                    {
                                        ulong function = BitConverter.ToUInt64(table, 15 * 8);
                                        if (function >= module.Start && function - module.Start < module.Size)
                                        {
                                            byte[] code = ReadBytes(process, function, 64);
                                            // 현재 버전의 GetUserClassID가 사용하는 RIP 상대 상수 읽기만 해석한다.
                                            if (code != null && code[0] == 0x0F && code[1] == 0x10 && code[2] == 0x05 && code[7] == 0x33 && code[8] == 0xC0 && code[9] == 0x0F && code[10] == 0x11 && code[11] == 0x02 && code[12] == 0xC3)
                                            {
                                                ulong constant = (ulong)((long)function + 7 + BitConverter.ToInt32(code, 3));
                                                byte[] guid = ReadBytes(process, constant, 16);
                                                if (guid != null)
                                                {
                                                    report.AppendFormat(" (GetUserClassID={0}, 객체=0x{1:X}, ID오프셋=0x{2:X})", new Guid(guid), begin + (ulong)offset, hit - begin - (ulong)offset);
                                                    if (new Guid(guid) == new Guid("459e8c62-070f-42b0-b92b-64dd59af0ebe")) objects.Add(begin + (ulong)offset);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (names.Contains(pointer)) report.AppendFormat("; 이름포인터[{0:+#;-#;0}]", (long)(begin + (ulong)offset) - (long)hit);
                        }
                        foreach (ulong name in names.Where(value => value >= begin && value < begin + (ulong)nearby.Length))
                            report.AppendFormat("; 이름[{0:+#;-#;0}]", (long)name - (long)hit);
                    }
                    report.Append("\r\n");
                }
                if (objects.Count > 0) TraceInput(process, input, objects, report);
                foreach (ulong candidate in objects.Distinct())
                {
                    byte[] value = ReadBytes(process, candidate, 0xE0);
                    if (value == null) continue;
                    // 현재 버전의 GetClientSite 구현을 대조하여 확인한 필드다.
                    ulong site = BitConverter.ToUInt64(value, 0x80);
                    report.AppendFormat("객체 0x{0:X}: 사이트(+80)=0x{1:X}\r\n", candidate, site);
                    byte[] siteValue = ReadBytes(process, site, 0x70);
                    if (siteValue != null)
                        report.AppendFormat("사이트 0x{0:X}: 소유자(+30)=0x{1:X}, 객체(+38)=0x{2:X}, 위치(+40)={3}, +68={4}\r\n", site,
                            BitConverter.ToUInt64(siteValue, 0x30), BitConverter.ToUInt64(siteValue, 0x38), BitConverter.ToInt32(siteValue, 0x40), BitConverter.ToInt32(siteValue, 0x68));
                }
                report.Append("ID 검색 후보에는 프로필 캐시·기존 메시지가 포함될 수 있다. 포인터 연결 후보도 삭제·복원 및 다른 사용자 비교로 확인해야 한다.\r\n");
                return report.ToString();
            }
            finally { CloseHandle(process); }
        }
        private static void TraceInput(IntPtr process, IntPtr input, List<ulong> objects, StringBuilder report)
        {
            var queue = new Queue<PointerNode>();
            var visited = new HashSet<ulong>();
            foreach (int index in new[] { 0, -21 })
            {
                ulong root = (ulong)GetWindowLongPtr(input, index).ToInt64();
                report.AppendFormat("입력창 포인터[{0}]=0x{1:X}\r\n", index, root);
                if (root != 0 && visited.Add(root)) queue.Enqueue(new PointerNode { Address = root, Path = "HWND[" + index + "]", Depth = 0 });
            }
            int readNodes = 0;
            while (queue.Count > 0 && readNodes < 4096)
            {
                var node = queue.Dequeue();
                MemoryInformation info;
                if (VirtualQueryEx(process, (IntPtr)(long)node.Address, out info, (UIntPtr)Marshal.SizeOf(typeof(MemoryInformation))) == UIntPtr.Zero ||
                    info.State != 0x1000 || info.Type != 0x20000 || (info.Protect & 0x101) != 0 || (info.Protect & 0xCC) == 0) continue;
                int length = (int)Math.Min(1024UL, (ulong)info.BaseAddress.ToInt64() + info.RegionSize.ToUInt64() - node.Address);
                byte[] bytes = ReadBytes(process, node.Address, length);
                if (bytes == null) continue;
                readNodes++;
                for (int offset = 0; offset + 8 <= bytes.Length; offset += 8)
                {
                    ulong pointer = BitConverter.ToUInt64(bytes, offset);
                    foreach (ulong target in objects)
                        if (pointer >= target && pointer - target <= 32 && (pointer - target) % 8 == 0)
                            report.AppendFormat("입력창 연결 후보: {0}+0x{1:X} -> 객체 0x{2:X}+0x{3:X}\r\n", node.Path, offset, target, pointer - target);
                    if (node.Depth < 4 && visited.Count < 32768 && pointer >= 0x10000 && pointer < 0x7FFFFFFFFFFF && pointer % 8 == 0 && visited.Add(pointer))
                        queue.Enqueue(new PointerNode { Address = pointer, Path = node.Path + "+0x" + offset.ToString("X") + " -> 0x" + pointer.ToString("X"), Depth = node.Depth + 1 });
                }
            }
            report.AppendFormat("입력 포인터 조사: {0}개 영역, 미처리 후보={1}\r\n", readNodes, queue.Count);
        }
        private sealed class PointerNode { internal ulong Address; internal int Depth; internal string Path; }
        private static void Find(byte[] bytes, byte[] pattern, ulong start, List<ulong> matches)
        {
            if (pattern.Length == 0 || matches.Count >= 512) return;
            int index = 0;
            while (index <= bytes.Length - pattern.Length && matches.Count < 512)
            {
                index = Array.IndexOf(bytes, pattern[0], index, bytes.Length - pattern.Length - index + 1);
                if (index < 0) break;
                int match = 1;
                while (match < pattern.Length && bytes[index + match] == pattern[match]) match++;
                if (match == pattern.Length && !matches.Contains(start + (ulong)index)) matches.Add(start + (ulong)index);
                index++;
            }
        }
        private static byte[] ReadBytes(IntPtr process, ulong address, int length)
        {
            byte[] bytes = new byte[length];
            UIntPtr read;
            return ReadProcessMemory(process, (IntPtr)(long)address, bytes, (UIntPtr)length, out read) && read.ToUInt64() == (ulong)length ? bytes : null;
        }
        private sealed class ModuleRange { internal ulong Start, Size; internal string Name; }
        [StructLayout(LayoutKind.Sequential)]
        private struct MemoryInformation
        {
            internal IntPtr BaseAddress, AllocationBase;
            internal uint AllocationProtect;
            internal ushort PartitionId;
            internal UIntPtr RegionSize;
            internal uint State, Protect, Type;
        }
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool IsWow64Process(IntPtr process, out bool wow64);
        [DllImport("kernel32.dll")] private static extern UIntPtr VirtualQueryEx(IntPtr process, IntPtr address, out MemoryInformation info, UIntPtr length);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, [Out] byte[] bytes, UIntPtr length, out UIntPtr read);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
    }
}
