using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace KakaotalkBot
{
    public class KakaoTalkDecryptor
    {
        private const int PageSize = 4096;
        private const int ReserveSize = 80;
        private const int ChunkSize = 1024 * 1024;
        private const uint ProcessAccess = 0x0410; // 프로세스 정보 조회 및 메모리 읽기 권한
        private static readonly byte[] Magic = Encoding.ASCII.GetBytes("SQLite format 3\0");
        private static readonly byte[] CodecSignature = Integers(256000, 2, 16, 32, 16, 16, PageSize);

        public sealed class Result
        {
            public string SourcePath { get; internal set; }
            public string OutputPath { get; internal set; }
            public string Status { get; internal set; }
            public string Error { get; internal set; }
            public int Pages { get; internal set; }
        }

        /// <summary>암호화된 DB를 확인한 시점의 메타데이터입니다.</summary>
        public sealed class DatabaseInfo
        {
            public string SourcePath { get; private set; }
            public int Pages { get; private set; }
            public byte[] Salt { get { return Sub(FirstPage, 0, 16); } }
            internal byte[] FirstPage { get; private set; }

            internal DatabaseInfo(string sourcePath, int pages, byte[] firstPage)
            {
                SourcePath = sourcePath;
                Pages = pages;
                FirstPage = firstPage;
            }
        }

        /// <summary>
        /// 특정 DB의 솔트에 대해 검증된 재사용 키입니다. 갱신 사이에도 보관하고 사용이 끝나면 해제합니다.
        /// 이 객체가 키 바이트를 소유하며, Dispose는 해당 키를 사용하는 작업이 끝날 때까지 기다립니다.
        /// </summary>
        public sealed class DecryptionKey : IDisposable
        {
            private readonly object sync = new object();
            private byte[] rawKey;
            private byte[] macKey;
            private byte[] salt;

            internal DecryptionKey(DatabaseInfo database, byte[] key)
            {
                rawKey = (byte[])key.Clone();
                salt = database.Salt;
                try
                {
                    macKey = MacKey(rawKey, salt);
                    DecryptPage(database.FirstPage, 1, rawKey, macKey);
                }
                catch
                {
                    Dispose();
                    throw;
                }
            }

            internal T Use<T>(Func<byte[], byte[], byte[], T> operation)
            {
                lock (sync)
                {
                    if (rawKey == null) throw new ObjectDisposedException("DecryptionKey");
                    return operation(rawKey, macKey, salt);
                }
            }

            public void Dispose()
            {
                lock (sync)
                {
                    if (rawKey != null) Array.Clear(rawKey, 0, rawKey.Length);
                    if (macKey != null) Array.Clear(macKey, 0, macKey.Length);
                    if (salt != null) Array.Clear(salt, 0, salt.Length);
                    rawKey = null;
                    macKey = null;
                    salt = null;
                }
            }
        }

        // 다른 PC에서도 호출자가 대상 카카오톡 계정의 폴더를 선택할 수 있도록 목록을 반환합니다.
        public static string[] FindChatDataDirectories()
        {
            string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kakao", "KakaoTalk", "users");
            return Directory.Exists(root)
                ? Directory.GetDirectories(root).Select(p => Path.Combine(p, "chat_data"))
                    .Where(Directory.Exists).ToArray()
                : new string[0];
        }

        public static int[] FindKakaoTalkProcessIds()
        {
            Process[] processes = Process.GetProcessesByName("KakaoTalk");
            try { return processes.Select(p => p.Id).ToArray(); }
            finally { foreach (Process process in processes) process.Dispose(); }
        }

        /// <summary>1단계: 호출자가 선택한 계정 폴더에서 채팅 DB 경로 목록을 조회합니다.</summary>
        public static string[] FindChatDatabaseFiles(string chatDataDirectory)
        {
            string directory = Path.GetFullPath(chatDataDirectory);
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException(directory);
            string[] files = Directory.GetFiles(directory, "chatLogs_*.edb");
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);
            return files;
        }

        /// <summary>2단계: 프로세스 메모리에 접근하지 않고 솔트와 첫 번째 암호화 페이지를 읽습니다.</summary>
        public DatabaseInfo ReadDatabaseInfo(string sourcePath)
        {
            string path = Path.GetFullPath(sourcePath);
            using (var source = OpenSource(path))
            {
                int pages = ReadPageCount(source);
                return new DatabaseInfo(path, pages, ReadExact(source, PageSize));
            }
        }

        /// <summary>
        /// 3단계: 검증된 키를 추출합니다. 현재 프로세스 메모리에서 키를 찾지 못하면 null을 반환합니다.
        /// 탐색 전에 카카오톡에서 대상 채팅방을 열어야 합니다. 반환된 키는 호출자가 소유하며 직접 해제해야 합니다.
        /// </summary>
        public DecryptionKey DiscoverKey(int processId, DatabaseInfo database, CancellationToken cancellation = default(CancellationToken))
        {
            if (database == null) throw new ArgumentNullException("database");
            IDictionary<string, DecryptionKey> keys = DiscoverKeys(processId, new[] { database }, cancellation);
            DecryptionKey key;
            return keys.TryGetValue(database.SourcePath, out key) ? key : null;
        }

        /// <summary>
        /// 한 번의 메모리 탐색으로 여러 DB의 키를 추출합니다. 반환 사전은 원본 파일의 전체 경로를 키로 사용합니다.
        /// 사전에 없는 DB는 키를 찾지 못한 것입니다. 호출자는 반환된 모든 키를 해제해야 합니다.
        /// </summary>
        public IDictionary<string, DecryptionKey> DiscoverKeys(int processId, IEnumerable<DatabaseInfo> databases,
            CancellationToken cancellation = default(CancellationToken))
        {
            Require64Bit();
            if (processId <= 0) throw new ArgumentOutOfRangeException("processId");
            if (databases == null) throw new ArgumentNullException("databases");
            var unique = new Dictionary<string, DatabaseInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (DatabaseInfo database in databases)
            {
                if (database == null) throw new ArgumentException("A database entry is null.", "databases");
                unique[database.SourcePath] = database;
            }
            if (unique.Count == 0) return new Dictionary<string, DecryptionKey>(StringComparer.OrdinalIgnoreCase);
            var targets = unique.Values.GroupBy(d => Convert.ToBase64String(d.Salt))
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);
            return DiscoverKeysCore(processId, targets, unique.Count, cancellation);
        }

        /// <summary>보유한 32바이트 원시 키를 DB에 대조하여 검증하고 재사용 가능한 키 객체를 생성합니다.</summary>
        public DecryptionKey CreateKey(DatabaseInfo database, byte[] rawKey)
        {
            if (database == null) throw new ArgumentNullException("database");
            if (rawKey == null || rawKey.Length != 32)
                throw new ArgumentException("A 32-byte raw key is required.", "rawKey");
            return new DecryptionKey(database, rawKey);
        }

        /// <summary>
        /// 4단계: 보관한 키로 전체 DB를 복호화하고 커밋된 WAL 프레임을 반영합니다.
        /// 프로세스 메모리를 다시 탐색하거나 키를 해제하지 않습니다. 변경된 페이지만 갱신하는 기능은 아닙니다.
        /// </summary>
        public int DecryptDatabase(string sourcePath, string outputPath, DecryptionKey key, bool overwrite = false)
        {
            if (key == null) throw new ArgumentNullException("key");
            return key.Use((rawKey, macKey, salt) =>
                DecryptFileCore(sourcePath, outputPath, rawKey, macKey, salt, overwrite));
        }

        /// <summary>
        /// 보관한 키로 4,096바이트 암호화 페이지 하나를 검증하고 복호화합니다. 페이지 번호는 1부터 시작합니다.
        /// 이 함수를 사용할 때 WAL 커밋 경계와 스냅샷의 일관성은 호출자가 처리해야 합니다.
        /// </summary>
        public byte[] DecryptPage(byte[] encryptedPage, int pageNumber, DecryptionKey key)
        {
            if (encryptedPage == null || encryptedPage.Length != PageSize)
                throw new ArgumentException("A 4096-byte encrypted page is required.", "encryptedPage");
            if (pageNumber < 1) throw new ArgumentOutOfRangeException("pageNumber");
            if (key == null) throw new ArgumentNullException("key");
            return key.Use((rawKey, macKey, salt) =>
            {
                if (pageNumber == 1) ValidateSalt(encryptedPage, salt);
                return DecryptPage(encryptedPage, pageNumber, rawKey, macKey);
            });
        }

        // 먼저 카카오톡에서 대상 채팅방을 열어야 합니다. 현재 프로세스 메모리에 있는 키만 찾을 수 있습니다.
        public IList<Result> DecryptAvailable(string chatDataDirectory, string outputDirectory, int processId,
            bool overwrite = false)
        {
            Require64Bit();
            string[] files = FindChatDatabaseFiles(chatDataDirectory);
            Directory.CreateDirectory(outputDirectory);
            var results = new List<Result>();
            var targets = new List<DatabaseInfo>();
            foreach (string file in files)
            {
                string output = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(file) + ".sqlite");
                if (File.Exists(output) && !overwrite)
                {
                    results.Add(new Result { SourcePath = file, OutputPath = output, Status = "SkippedExisting" });
                    continue;
                }
                try { targets.Add(ReadDatabaseInfo(file)); }
                catch (Exception ex)
                {
                    if (!(ex is IOException || ex is InvalidDataException || ex is UnauthorizedAccessException)) throw;
                    results.Add(new Result { SourcePath = file, OutputPath = output, Status = "Failed", Error = ex.Message });
                }
            }
            if (targets.Count == 0) return results;
            IDictionary<string, DecryptionKey> keys = DiscoverKeys(processId, targets);
            try
            {
                foreach (DatabaseInfo target in targets)
                {
                    string output = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(target.SourcePath) + ".sqlite");
                    var result = new Result { SourcePath = target.SourcePath, OutputPath = output };
                    DecryptionKey key;
                    if (!keys.TryGetValue(target.SourcePath, out key)) result.Status = "KeyUnavailable";
                    else
                    {
                        try
                        {
                            result.Pages = DecryptDatabase(target.SourcePath, output, key, overwrite);
                            result.Status = "Decrypted";
                        }
                        catch (Exception ex)
                        {
                            result.Status = "Failed";
                            result.Error = ex.Message;
                        }
                    }
                    results.Add(result);
                }
            }
            finally { foreach (DecryptionKey key in keys.Values) key.Dispose(); }
            return results;
        }

        // 호출자가 이미 추출한 키를 사용하는 함수입니다. 원시 키의 길이는 정확히 32바이트여야 합니다.
        public int DecryptFile(string sourcePath, string outputPath, byte[] key, bool overwrite = false)
        {
            if (key == null || key.Length != 32) throw new ArgumentException("A 32-byte raw key is required.", "key");
            using (DecryptionKey reusableKey = CreateKey(ReadDatabaseInfo(sourcePath), key))
                return DecryptDatabase(sourcePath, outputPath, reusableKey, overwrite);
        }

        private static int DecryptFileCore(string sourcePath, string outputPath, byte[] key, byte[] macKey,
            byte[] salt, bool overwrite)
        {
            sourcePath = Path.GetFullPath(sourcePath);
            outputPath = Path.GetFullPath(outputPath);
            if (new[] { sourcePath, sourcePath + "-wal", sourcePath + "-shm" }
                .Any(path => string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException("Output must not replace the source database or its sidecar files.", "outputPath");
            if (File.Exists(outputPath) && !overwrite) throw new IOException("Output already exists: " + outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
            string temporary = outputPath + "." + Guid.NewGuid().ToString("N") + ".partial";
            try
            {
                int pages;
                using (var source = OpenSource(sourcePath))
                {
                    pages = ReadPageCount(source);
                    byte[] first = ReadExact(source, PageSize);
                    ValidateSalt(first, salt);
                    source.Position = 0;
                    using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite))
                    {
                        for (int number = 1; number <= pages; number++)
                        {
                            byte[] page = DecryptPage(ReadExact(source, PageSize), number, key, macKey);
                            output.Write(page, 0, page.Length);
                        }
                        ApplyWal(sourcePath + "-wal", output, key, macKey);
                        output.Flush(true);
                    }
                }
                ValidateHeader(temporary);
                if (overwrite && File.Exists(outputPath)) File.Replace(temporary, outputPath, null);
                else File.Move(temporary, outputPath);
                return pages;
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static FileStream OpenSource(string path)
        {
            return File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        }

        private static int ReadPageCount(FileStream source)
        {
            if (source.Length < PageSize || source.Length % PageSize != 0)
                throw new InvalidDataException("Source is not a whole number of 4096-byte pages.");
            return checked((int)(source.Length / PageSize));
        }

        private static void ValidateSalt(byte[] firstPage, byte[] salt)
        {
            if (!Equal(firstPage, 0, salt, 0, salt.Length))
                throw new CryptographicException("Database salt changed or the key belongs to another database. Read database info and discover the key again.");
        }

        private static void ApplyWal(string walPath, FileStream output, byte[] key, byte[] macKey)
        {
            if (!File.Exists(walPath)) return;
            using (var wal = File.Open(walPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                if (wal.Length <= 32) return;
                byte[] header = ReadExact(wal, 32);
                if (!(header[0] == 0x37 && header[1] == 0x7f && header[2] == 0x06 &&
                      (header[3] == 0x82 || header[3] == 0x83)))
                    throw new InvalidDataException("Unrecognized WAL magic.");
                if (BigEndian32(header, 8) != PageSize) throw new InvalidDataException("Unexpected WAL page size.");
                long frames = (wal.Length - 32) / (24 + PageSize);
                long committed = 0;
                uint committedSize = 0;
                for (long i = 0; i < frames; i++)
                {
                    wal.Position = 32 + i * (24 + PageSize);
                    byte[] frameHeader = ReadExact(wal, 24);
                    if (!Equal(frameHeader, 8, header, 16, 8)) break;
                    uint size = BigEndian32(frameHeader, 4);
                    if (size != 0) { committed = i + 1; committedSize = size; }
                }
                for (long i = 0; i < committed; i++)
                {
                    wal.Position = 32 + i * (24 + PageSize);
                    byte[] frameHeader = ReadExact(wal, 24);
                    uint number = BigEndian32(frameHeader, 0);
                    if (number == 0 || number > int.MaxValue) throw new InvalidDataException("Invalid WAL page number.");
                    byte[] page = DecryptPage(ReadExact(wal, PageSize), (int)number, key, macKey);
                    output.Position = ((long)number - 1) * PageSize;
                    output.Write(page, 0, page.Length);
                }
                if (committed != 0) output.SetLength((long)committedSize * PageSize);
            }
        }

        private static byte[] DecryptPage(byte[] page, int number, byte[] key, byte[] macKey)
        {
            int start = number == 1 ? 16 : 0;
            int length = PageSize - ReserveSize - start;
            byte[] message = Sub(page, start, length + 16);
            byte[] counter = BitConverter.GetBytes((uint)number);
            byte[] authenticated = new byte[message.Length + 4];
            Buffer.BlockCopy(message, 0, authenticated, 0, message.Length);
            Buffer.BlockCopy(counter, 0, authenticated, message.Length, 4);
            using (var hmac = new HMACSHA512(macKey))
                if (!FixedEqual(hmac.ComputeHash(authenticated), Sub(page, PageSize - 64, 64)))
                    throw new CryptographicException("HMAC failed on page " + number);
            byte[] plain;
            using (var aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = Sub(page, PageSize - ReserveSize, 16);
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                    plain = decryptor.TransformFinalBlock(page, start, length);
            }
            byte[] output = new byte[PageSize];
            if (number == 1) Buffer.BlockCopy(Magic, 0, output, 0, Magic.Length);
            Buffer.BlockCopy(plain, 0, output, start, plain.Length);
            return output;
        }

        private Dictionary<string, DecryptionKey> DiscoverKeysCore(int pid,
            Dictionary<string, List<DatabaseInfo>> targets, int targetCount, CancellationToken cancellation)
        {
            IntPtr process = OpenProcess(ProcessAccess, false, pid);
            if (process == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(),
                "Cannot read KakaoTalk process. Run as the same Windows user and elevation level.");
            var keys = new Dictionary<string, DecryptionKey>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (MemoryChunk chunk in MemoryChunks(process))
                {
                    cancellation.ThrowIfCancellationRequested();
                    int at = 0;
                    while ((at = IndexOf(chunk.Bytes, CodecSignature, at)) >= 0)
                    {
                        ulong baseAddress = chunk.Address + (ulong)at - 4;
                        byte[] context = ReadMemory(process, baseAddress, 128);
                        if (context != null && BitConverter.ToUInt32(context, 12) == 16 &&
                            BitConverter.ToUInt32(context, 16) == 32 &&
                            BitConverter.ToUInt32(context, 20) == 16 &&
                            BitConverter.ToUInt32(context, 28) == PageSize &&
                            BitConverter.ToUInt32(context, 36) == ReserveSize &&
                            BitConverter.ToUInt32(context, 40) == 64)
                        {
                            byte[] salt = ReadMemory(process, BitConverter.ToUInt64(context, 64), 16);
                            List<DatabaseInfo> matches;
                            if (salt != null && targets.TryGetValue(Convert.ToBase64String(salt), out matches) &&
                                matches.Any(target => !keys.ContainsKey(target.SourcePath)))
                            {
                                foreach (int offset in new[] { 96, 104 })
                                {
                                    byte[] cipherContext = ReadMemory(process, BitConverter.ToUInt64(context, offset), 32);
                                    if (cipherContext == null) continue;
                                    byte[] key = ReadMemory(process, BitConverter.ToUInt64(cipherContext, 8), 32);
                                    if (key == null) continue;
                                    try
                                    {
                                        foreach (DatabaseInfo target in matches)
                                        {
                                            if (keys.ContainsKey(target.SourcePath)) continue;
                                            try { keys.Add(target.SourcePath, CreateKey(target, key)); }
                                            catch (CryptographicException) { /* 다음 후보 키로 검증을 시도합니다. */ }
                                        }
                                    }
                                    finally { Array.Clear(key, 0, key.Length); }
                                    if (matches.All(target => keys.ContainsKey(target.SourcePath))) break;
                                }
                            }
                        }
                        at++;
                    }
                    if (keys.Count == targetCount) break;
                }
            }
            catch
            {
                foreach (DecryptionKey key in keys.Values) key.Dispose();
                throw;
            }
            finally { CloseHandle(process); }
            return keys;
        }

        private static byte[] MacKey(byte[] key, byte[] salt)
        {
            byte[] altered = salt.Select(b => (byte)(b ^ 0x3a)).ToArray();
            using (var hmac = new HMACSHA512(key))
            {
                byte[] first = new byte[altered.Length + 4];
                Buffer.BlockCopy(altered, 0, first, 0, altered.Length);
                first[first.Length - 1] = 1;
                byte[] u1 = hmac.ComputeHash(first);
                byte[] u2 = hmac.ComputeHash(u1);
                return Enumerable.Range(0, 32).Select(i => (byte)(u1[i] ^ u2[i])).ToArray();
            }
        }

        private static void ValidateHeader(string path)
        {
            using (var file = File.OpenRead(path))
            {
                byte[] header = ReadExact(file, 100);
                if (!Equal(header, 0, Magic, 0, Magic.Length) || BigEndian32(header, 16) >> 16 != PageSize)
                    throw new InvalidDataException("Decrypted SQLite header is invalid.");
            }
        }

        private static IEnumerable<MemoryChunk> MemoryChunks(IntPtr process)
        {
            ulong address = 0;
            ulong maximum = 0x00007FFFFFFFFFFFUL;
            while (address < maximum)
            {
                MEMORY_BASIC_INFORMATION info;
                if (VirtualQueryEx(process, new IntPtr(unchecked((long)address)), out info,
                    (UIntPtr)Marshal.SizeOf(typeof(MEMORY_BASIC_INFORMATION))) == UIntPtr.Zero) yield break;
                ulong baseAddress = unchecked((ulong)info.BaseAddress.ToInt64());
                ulong size = info.RegionSize.ToUInt64();
                if (info.State == 0x1000 && (info.Protect & 0x101) == 0 &&
                    (info.Protect & 0xEE) != 0)
                {
                    // 청크 경계에 걸친 시그니처도 찾으려면 1바이트보다 넓은 범위를 겹쳐 읽어야 합니다.
                    for (ulong offset = 0; offset < size; offset += ChunkSize)
                    {
                        int length = (int)Math.Min((ulong)(ChunkSize + CodecSignature.Length - 1), size - offset);
                        byte[] bytes = ReadMemory(process, baseAddress + offset, length);
                        if (bytes != null) yield return new MemoryChunk { Address = baseAddress + offset, Bytes = bytes };
                    }
                }
                ulong next = baseAddress + size;
                if (next <= address) yield break;
                address = next;
            }
        }

        private static byte[] ReadMemory(IntPtr process, ulong address, int length)
        {
            if (address == 0 || length <= 0) return null;
            byte[] bytes = new byte[length];
            UIntPtr read;
            return ReadProcessMemory(process, new IntPtr(unchecked((long)address)), bytes,
                (UIntPtr)length, out read) && read.ToUInt64() == (ulong)length ? bytes : null;
        }

        private static void Require64Bit()
        {
            if (!Environment.Is64BitProcess || !BitConverter.IsLittleEndian)
                throw new PlatformNotSupportedException("Build and run this class as x64 on Windows.");
        }

        private static int IndexOf(byte[] data, byte[] pattern, int from)
        {
            for (int i = from; i <= data.Length - pattern.Length; i++)
            {
                int j = 0;
                while (j < pattern.Length && data[i + j] == pattern[j]) j++;
                if (j == pattern.Length) return i;
            }
            return -1;
        }

        private static byte[] ReadExact(Stream stream, int count)
        {
            byte[] result = new byte[count];
            int offset = 0;
            while (offset < count)
            {
                int n = stream.Read(result, offset, count - offset);
                if (n == 0) throw new EndOfStreamException();
                offset += n;
            }
            return result;
        }

        private static byte[] Sub(byte[] source, int offset, int length)
        {
            byte[] result = new byte[length];
            Buffer.BlockCopy(source, offset, result, 0, length);
            return result;
        }

        private static bool Equal(byte[] a, int aOffset, byte[] b, int bOffset, int length)
        {
            for (int i = 0; i < length; i++) if (a[aOffset + i] != b[bOffset + i]) return false;
            return true;
        }

        private static bool FixedEqual(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            int diff = 0;
            for (int i = 0; i < a.Length; i++) diff |= a[i] ^ b[i];
            return diff == 0;
        }

        private static uint BigEndian32(byte[] data, int offset)
        {
            return ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) |
                ((uint)data[offset + 2] << 8) | data[offset + 3];
        }

        private static byte[] Integers(params int[] values)
        {
            return values.SelectMany(BitConverter.GetBytes).ToArray();
        }

        private sealed class MemoryChunk { public ulong Address; public byte[] Bytes; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MEMORY_BASIC_INFORMATION
        {
            public IntPtr BaseAddress;
            public IntPtr AllocationBase;
            public uint AllocationProtect;
            public ushort PartitionId;
            public UIntPtr RegionSize;
            public uint State;
            public uint Protect;
            public uint Type;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint access, bool inheritHandle, int processId);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, [Out] byte[] buffer,
            UIntPtr size, out UIntPtr bytesRead);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern UIntPtr VirtualQueryEx(IntPtr process, IntPtr address,
            out MEMORY_BASIC_INFORMATION info, UIntPtr length);
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);
    }
}
