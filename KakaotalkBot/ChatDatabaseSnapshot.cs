using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

namespace KakaotalkBot
{
    // 한 작업자만 소유합니다. 원본에는 잠금을 걸지 않고 읽기 전후 상태를 비교합니다.
    // SQLite 잠금/체크섬 규격: https://www.sqlite.org/walformat.html
    internal sealed class ChatDatabaseSnapshot
    {
        private const int PageSize = 4096;
        private const int FrameSize = PageSize + 24;
        private readonly string sourcePath;
        private readonly KakaoTalkDecryptor decryptor;
        private readonly KakaoTalkDecryptor.DecryptionKey key;
        private readonly Dictionary<int, byte[]> hashes = new Dictionary<int, byte[]>();
        // 수신 세션은 순차 실행됩니다. 큰 사본 버퍼는 세션 내에서 재사용합니다.
        private byte[] copyBuffer;
        private byte[] walHeader;
        private uint lastFrame, checksum1, checksum2;
        private string baseStamp;
        private bool valid;
        public string OutputPath { get; private set; }
        public int LastDecryptedPages { get; private set; }
        // 원본이 바뀐 주기에는 사본과 메시지 읽기 기준점을 유지합니다.
        public bool RefreshDeferred { get; private set; }

        public ChatDatabaseSnapshot(string source, string output, KakaoTalkDecryptor decryptor,
            KakaoTalkDecryptor.DecryptionKey key)
        {
            sourcePath = Path.GetFullPath(source);
            OutputPath = Path.GetFullPath(output);
            if (OutputPath.StartsWith(sourcePath, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("복호화 사본은 원본과 다른 경로에 저장해야 합니다.");
            this.decryptor = decryptor;
            this.key = key;
            Directory.CreateDirectory(Path.GetDirectoryName(OutputPath));
        }

        public bool Refresh(CancellationToken cancellation)
        {
            RefreshDeferred = false;
            using (Capture capture = ReadCapture(cancellation))
            {
                LastDecryptedPages = 0;
                if (capture == null) return false;
                bool rebuild = !valid;
                string target = rebuild ? OutputPath + ".building" : OutputPath;
                var updates = new Dictionary<int, byte[]>();
                var newHashes = new Dictionary<int, byte[]>();
                using (SHA256 sha = SHA256.Create())
                {
                    Action<int, byte[]> inspect = (number, encrypted) =>
                    {
                        cancellation.ThrowIfCancellationRequested();
                        byte[] hash = sha.ComputeHash(encrypted);
                        byte[] previous;
                        if (!rebuild && hashes.TryGetValue(number, out previous) && previous.SequenceEqual(hash)) return;
                        updates[number] = decryptor.DecryptPage(encrypted, number, key);
                        newHashes[number] = hash;
                    };
                    if (capture.BaseCopy != null)
                    {
                        using (var input = File.OpenRead(capture.BaseCopy))
                        {
                            for (int number = 1; number <= capture.Pages; number++)
                            {
                                byte[] frame;
                                if (capture.Frames.TryGetValue(number, out frame)) { inspect(number, frame); continue; }
                                if ((long)number * PageSize > input.Length)
                                    throw new InvalidDataException("DB 확장 페이지가 WAL에 없습니다.");
                                input.Position = ((long)number - 1) * PageSize;
                                inspect(number, Read(input, PageSize));
                            }
                        }
                    }
                    else foreach (var frame in capture.Frames.Where(f => f.Key <= capture.Pages)) inspect(frame.Key, frame.Value);
                }

                // 원본 읽기와 복호화가 모두 성공한 뒤 사본을 갱신합니다. 조회 연결은 이 시점에 닫혀 있습니다.
                var undo = new Dictionary<int, byte[]>();
                long oldLength = 0;
                bool writesStarted = false;
                try
                {
                    using (var output = new FileStream(target, rebuild ? FileMode.Create : FileMode.Open,
                        FileAccess.ReadWrite, FileShare.None))
                    {
                        oldLength = output.Length;
                        if (!rebuild)
                        {
                            foreach (int number in updates.Keys)
                            {
                                long position = ((long)number - 1) * PageSize;
                                if (position < oldLength) { output.Position = position; undo[number] = Read(output, PageSize); }
                            }
                            for (long position = (long)capture.Pages * PageSize; position < oldLength; position += PageSize)
                            { output.Position = position; undo[checked((int)(position / PageSize + 1))] = Read(output, PageSize); }
                        }
                        writesStarted = true;
                        foreach (var update in updates)
                        {
                            cancellation.ThrowIfCancellationRequested();
                            output.Position = ((long)update.Key - 1) * PageSize;
                            output.Write(update.Value, 0, PageSize);
                        }
                        output.SetLength((long)capture.Pages * PageSize);
                        // WAL을 합친 독립 사본이므로 SQLite가 새로운 WAL을 찾지 않도록 파일 모드를 지정합니다.
                        output.Position = 18;
                        output.WriteByte(1);
                        output.WriteByte(1);
                        output.Flush(true);
                    }
                    using (var db = new ChatSqlite(target))
                    {
                        db.Scalar("SELECT count(*) FROM sqlite_master");
                        if (rebuild || capture.BaseCopy != null)
                        {
                            var check = db.Query("PRAGMA quick_check");
                            if (check.Count != 1 || Convert.ToString(check[0][0]) != "ok")
                                throw new InvalidDataException("복호화한 DB의 무결성 검사에 실패했습니다.");
                        }
                    }
                    if (rebuild)
                    {
                        if (File.Exists(OutputPath)) File.Replace(target, OutputPath, null);
                        else File.Move(target, OutputPath);
                        hashes.Clear();
                    }
                    foreach (var item in newHashes) hashes[item.Key] = item.Value;
                    foreach (int number in hashes.Keys.Where(n => n > capture.Pages).ToArray()) hashes.Remove(number);
                    walHeader = capture.Header;
                    lastFrame = capture.Committed;
                    checksum1 = capture.Checksum1;
                    checksum2 = capture.Checksum2;
                    baseStamp = capture.Stamp;
                    valid = true;
                    LastDecryptedPages = updates.Count;
                    return true;
                }
                catch
                {
                    if (!rebuild && writesStarted)
                    {
                        try
                        {
                            using (var output = new FileStream(OutputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                            {
                                output.SetLength(oldLength);
                                foreach (var page in undo)
                                { output.Position = ((long)page.Key - 1) * PageSize; output.Write(page.Value, 0, PageSize); }
                                output.Flush(true);
                            }
                        }
                        catch { valid = false; }
                    }
                    throw;
                }
                finally { if (rebuild && File.Exists(target)) File.Delete(target); }
            }
        }

        private Capture ReadCapture(CancellationToken cancellation)
        {
            var result = new Capture();
            try
            {
                using (var reader = new SourceReader(sourcePath))
                {
                    cancellation.ThrowIfCancellationRequested();
                    var source = reader.Source;
                    if (source.Length < PageSize || source.Length % PageSize != 0)
                        throw new InvalidDataException("암호화 DB 페이지 크기가 잘못되었습니다.");
                    // 첫 페이지 HMAC를 확인하면 동일 경로의 DB가 교체된 경우도 이전 키로 진행하지 않습니다.
                    byte[] first = Read(source, PageSize);
                    decryptor.DecryptPage(first, 1, key);
                    result.Stamp = source.Length + ":" + File.GetLastWriteTimeUtc(sourcePath).Ticks + ":" + Convert.ToBase64String(first.Skip(4016).ToArray());
                    result.Pages = checked((int)(source.Length / PageSize));
                    bool append = false;
                    string walPath = sourcePath + "-wal";
                    if (File.Exists(walPath))
                    using (var wal = new FileStream(walPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    {
                        if (wal.Length >= 32)
                        {
                            result.Header = Read(wal, 32);
                            uint magic = Big(result.Header, 0);
                            if ((magic != 0x377f0682 && magic != 0x377f0683) || Big(result.Header, 8) != PageSize || Big(result.Header, 4) != 3007000)
                                throw new InvalidDataException("지원하지 않는 WAL 헤더입니다.");
                            bool big = magic == 0x377f0683;
                            uint s1 = 0, s2 = 0;
                            Sum(result.Header, 0, 24, big, ref s1, ref s2);
                            if (s1 != Big(result.Header, 24) || s2 != Big(result.Header, 28)) throw new InvalidDataException("WAL 헤더 체크섬 오류입니다.");
                            uint maximum = checked((uint)((wal.Length - 32) / FrameSize));
                            if (reader.Index != null)
                            {
                                if (!reader.Index.Take(48).SequenceEqual(reader.Index.Skip(48).Take(48)) || reader.Index[12] != 1)
                                    throw new IOException("WAL 인덱스 준비 중입니다.");
                                uint i1 = 0, i2 = 0;
                                Sum(reader.Index, 0, 40, false, ref i1, ref i2);
                                if (i1 != BitConverter.ToUInt32(reader.Index, 40) || i2 != BitConverter.ToUInt32(reader.Index, 44))
                                    throw new IOException("WAL 인덱스 체크섬 오류입니다.");
                                uint published = BitConverter.ToUInt32(reader.Index, 16);
                                if (published > maximum) throw new IOException("WAL의 커밋 프레임을 아직 읽을 수 없습니다.");
                                maximum = published;
                                if (maximum > 0 && !reader.Index.Skip(32).Take(8).SequenceEqual(result.Header.Skip(16).Take(8)))
                                    throw new IOException("WAL 세대가 변경되었습니다.");
                            }
                            append = valid && walHeader != null && walHeader.SequenceEqual(result.Header) && maximum >= lastFrame;
                            if (append && lastFrame > 0)
                            {
                                wal.Position = 32 + ((long)lastFrame - 1) * FrameSize;
                                byte[] old = Read(wal, 24);
                                append = Big(old, 16) == checksum1 && Big(old, 20) == checksum2;
                            }
                            uint start = append ? lastFrame : 0;
                            if (append) { s1 = checksum1; s2 = checksum2; result.Committed = lastFrame; }
                            result.Checksum1 = s1;
                            result.Checksum2 = s2;
                            var pending = new Dictionary<int, byte[]>();
                            for (uint i = start; i < maximum; i++)
                            {
                                cancellation.ThrowIfCancellationRequested();
                                wal.Position = 32 + (long)i * FrameSize;
                                byte[] header = Read(wal, 24);
                                byte[] page = Read(wal, PageSize);
                                Sum(header, 0, 8, big, ref s1, ref s2);
                                Sum(page, 0, PageSize, big, ref s1, ref s2);
                                bool correct = header.Skip(8).Take(8).SequenceEqual(result.Header.Skip(16).Take(8)) && s1 == Big(header, 16) && s2 == Big(header, 20);
                                if (!correct)
                                {
                                    if (reader.Index != null) throw new IOException("커밋된 WAL 프레임 검증에 실패했습니다.");
                                    break;
                                }
                                int number = checked((int)Big(header, 0));
                                if (number < 1) throw new InvalidDataException("WAL 페이지 번호 오류입니다.");
                                pending[number] = page;
                                uint size = Big(header, 4);
                                if (size != 0)
                                {
                                    foreach (var item in pending) result.Frames[item.Key] = item.Value;
                                    pending.Clear();
                                    result.Pages = checked((int)size);
                                    result.Committed = i + 1;
                                    result.Checksum1 = s1;
                                    result.Checksum2 = s2;
                                }
                            }
                            if (reader.Index != null && result.Committed != maximum) throw new IOException("WAL의 마지막 커밋을 확인하지 못했습니다.");
                            if (append && result.Committed == lastFrame) { RefreshDeferred = !reader.Verify(); return null; }
                        }
                    }
                    if (result.Header == null && valid && walHeader == null && result.Stamp == baseStamp) { RefreshDeferred = !reader.Verify(); return null; }
                    if (!append)
                    {
                        result.BaseCopy = OutputPath + ".capture";
                        source.Position = 0;
                        var watch = Stopwatch.StartNew();
                        using (var copy = new FileStream(result.BaseCopy, FileMode.Create, FileAccess.Write))
                        {
                            if (copyBuffer == null) copyBuffer = new byte[1024 * 1024];
                            byte[] buffer = copyBuffer;
                            int count;
                            while ((count = source.Read(buffer, 0, buffer.Length)) != 0)
                            {
                                cancellation.ThrowIfCancellationRequested();
                                if (watch.ElapsedMilliseconds > 2000) throw new IOException("DB 사본 읽기가 오래 걸려 중단했습니다. 다시 시도합니다.");
                                copy.Write(buffer, 0, count);
                            }
                        }
                        // 체크포인트 도중 여러 시점의 페이지가 섞이지 않았는지 원본을 다시 읽어 확인합니다.
                        source.Position = 0;
                        using (SHA256 sha = SHA256.Create())
                        using (var copy = File.OpenRead(result.BaseCopy))
                        {
                            byte[] copied = sha.ComputeHash(copy);
                            cancellation.ThrowIfCancellationRequested();
                            if (!copied.SequenceEqual(sha.ComputeHash(source)))
                                { RefreshDeferred = true; return null; }
                        }
                    }
                    if (!reader.Verify()) { RefreshDeferred = true; return null; }
                }
                return result;
            }
            catch { result.Dispose(); throw; }
            finally { if (RefreshDeferred) result.Dispose(); }
        }

        private sealed class Capture : IDisposable
        {
            public string BaseCopy, Stamp;
            public byte[] Header;
            public uint Committed, Checksum1, Checksum2;
            public int Pages;
            public readonly Dictionary<int, byte[]> Frames = new Dictionary<int, byte[]>();
            public void Dispose() { if (BaseCopy != null && File.Exists(BaseCopy)) File.Delete(BaseCopy); }
        }

        internal static uint Big(byte[] bytes, int offset)
        { return ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3]; }

        internal static void Sum(byte[] bytes, int offset, int count, bool big, ref uint s1, ref uint s2)
        {
            unchecked
            {
                for (int i = offset; i < offset + count; i += 8)
                {
                    s1 += (big ? Big(bytes, i) : BitConverter.ToUInt32(bytes, i)) + s2;
                    s2 += (big ? Big(bytes, i + 4) : BitConverter.ToUInt32(bytes, i + 4)) + s1;
                }
            }
        }

        internal static byte[] Read(Stream stream, int count)
        {
            byte[] result = new byte[count];
            int offset = 0;
            while (offset < count)
            { int read = stream.Read(result, offset, count - offset); if (read == 0) throw new EndOfStreamException(); offset += read; }
            return result;
        }

        // 원본은 공유 읽기로만 엽니다. SQLite의 잠금 영역에는 접근하지 않습니다.
        // 상태가 바뀌면 사본을 폐기하고 다음 수신 주기에 다시 시도합니다.
        private sealed class SourceReader : IDisposable
        {
            public FileStream Source;
            public byte[] Index;
            private readonly string path;
            private byte[] state;
            public SourceReader(string path)
            {
                this.path = path;
                try
                {
                    Source = Open(path);
                    state = ReadState();
                }
                catch { Dispose(); throw; }
            }
            private static FileStream Open(string file)
            {
                // 버퍼의 선행 읽기가 SHM 잠금 영역까지 넘어가지 않도록 버퍼링을 끕니다.
                return new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1);
            }

            private byte[] ReadState()
            {
                using (var stateStream = new MemoryStream())
                using (var writer = new BinaryWriter(stateStream))
                {
                    // 새 핸들로 읽어 경로의 파일 교체와 FileStream 읽기 버퍼의 영향을 피합니다.
                    using (var db = Open(path))
                    {
                        writer.Write(db.Length);
                        writer.Write(File.GetLastWriteTimeUtc(path).Ticks);
                        writer.Write(Read(db, PageSize));
                    }
                    bool hasShm = File.Exists(path + "-shm");
                    writer.Write(hasShm);
                    if (hasShm)
                    using (var shm = Open(path + "-shm"))
                    {
                        byte[] header = Read(shm, 96);
                        writer.Write(header);
                        if (Index == null) Index = header;
                        // 완료/시도한 체크포인트 위치를 비교하되 잠금 바이트 120~127은 읽지 않습니다.
                        writer.Write(Read(shm, 4));
                        shm.Position = 128;
                        writer.Write(Read(shm, 8));
                    }
                    bool hasWal = File.Exists(path + "-wal");
                    writer.Write(hasWal);
                    if (hasWal)
                    using (var wal = Open(path + "-wal"))
                    {
                        writer.Write(wal.Length);
                        writer.Write(Read(wal, (int)Math.Min(32, wal.Length)));
                    }
                    if (File.Exists(path + "-journal"))
                    using (var journal = Open(path + "-journal"))
                    {
                        if (journal.Length > 0 && Read(journal, (int)Math.Min(8, journal.Length)).Any(b => b != 0))
                            throw new IOException("롤백 저널이 남아 있습니다. 카카오톡의 복구 완료 후 다시 시도합니다.");
                    }
                    return stateStream.ToArray();
                }
            }
            public bool Verify()
            {
                return state.SequenceEqual(ReadState());
            }
            public void Dispose()
            {
                if (Source != null) Source.Dispose();
            }
        }
    }
}
