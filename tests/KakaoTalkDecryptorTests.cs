using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using KakaotalkBot;

// Standalone regression tests. Uses generated encrypted pages and this test process only.
internal static class KakaoTalkDecryptorTests
{
    private const int PageSize = 4096;
    private static int passed;

    private static int Main()
    {
        string directory = Path.Combine(Path.GetTempPath(), "KakaoDecryptorTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            Run(directory);
            Console.WriteLine("PASS: " + passed + " checks (synthetic databases and process memory)");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
        finally { Directory.Delete(directory, true); }
    }

    private static void Run(string directory)
    {
        var decryptor = new KakaoTalkDecryptor();
        byte[] rawKey = Enumerable.Range(1, 32).Select(i => (byte)i).ToArray();
        byte[] salt = Enumerable.Range(51, 16).Select(i => (byte)i).ToArray();
        byte[] first = PlainPage(1, 21);
        byte[] second = PlainPage(2, 22);
        byte[] encryptedFirst = EncryptPage(first, 1, rawKey, salt);
        byte[] encryptedSecond = EncryptPage(second, 2, rawKey, salt);
        string source = Path.Combine(directory, "chatLogs_123.edb");
        string copy = Path.Combine(directory, "chatLogs_124.edb");
        string output = Path.Combine(directory, "result.sqlite");
        File.WriteAllBytes(source, encryptedFirst.Concat(encryptedSecond).ToArray());
        File.Copy(source, copy);

        Check(KakaoTalkDecryptor.FindChatDatabaseFiles(directory).Length == 2, "enumerate databases");
        var info = decryptor.ReadDatabaseInfo(source);
        Check(info.SourcePath == Path.GetFullPath(source) && info.Pages == 2, "metadata");
        byte[] publicSalt = info.Salt;
        publicSalt[0] ^= 0xff;
        Check(info.Salt.SequenceEqual(salt), "salt copy cannot mutate database info");
        Expect<CryptographicException>(() => decryptor.CreateKey(info, new byte[32]), "reject wrong key");
        Expect<ArgumentException>(() => decryptor.CreateKey(info, new byte[31]), "reject wrong key length");

        byte[] callerKey = (byte[])rawKey.Clone();
        var key = decryptor.CreateKey(info, callerKey);
        try
        {
            Array.Clear(callerKey, 0, callerKey.Length);
            Check(decryptor.DecryptPage(encryptedFirst, 1, key).SequenceEqual(first), "key owns a copy; page 1");
            Check(decryptor.DecryptPage(encryptedSecond, 2, key).SequenceEqual(second), "page 2");
            Expect<CryptographicException>(() => decryptor.DecryptPage(encryptedSecond, 3, key), "authenticated page number");
            byte[] damaged = (byte[])encryptedSecond.Clone();
            damaged[100] ^= 1;
            Expect<CryptographicException>(() => decryptor.DecryptPage(damaged, 2, key), "reject corrupt page");
            Expect<ArgumentOutOfRangeException>(() => decryptor.DecryptPage(encryptedFirst, 0, key), "reject page zero");
            Expect<ArgumentException>(() => decryptor.DecryptPage(new byte[10], 1, key), "reject short page");

            Check(decryptor.DecryptDatabase(source, output, key) == 2, "full decrypt page count");
            byte[] expected = first.Concat(second).ToArray();
            Check(File.ReadAllBytes(output).SequenceEqual(expected), "full plaintext matches fixture");
            Expect<IOException>(() => decryptor.DecryptDatabase(source, output, key), "protect existing output");
            decryptor.DecryptDatabase(source, output, key, true);
            Check(File.ReadAllBytes(output).SequenceEqual(expected), "reuse key for refresh");
            decryptor.DecryptDatabase(copy, output, key, true);
            Check(File.ReadAllBytes(output).SequenceEqual(expected), "reuse key for same database copy");
            Expect<ArgumentException>(() => decryptor.DecryptDatabase(source, source, key, true), "protect source");
            Expect<ArgumentException>(() => decryptor.DecryptDatabase(source, source + "-wal", key, true), "protect WAL");
            Expect<ArgumentException>(() => decryptor.DecryptDatabase(source, source + "-shm", key, true), "protect SHM");

            File.WriteAllBytes(source, encryptedFirst.Concat(damaged).ToArray());
            Expect<CryptographicException>(() => decryptor.DecryptDatabase(source, output, key, true), "failed refresh");
            Check(File.ReadAllBytes(output).SequenceEqual(expected), "failure preserves previous output");
            Check(Directory.GetFiles(directory, "*.partial").Length == 0, "partial output cleaned");

            byte[] otherSalt = (byte[])salt.Clone();
            otherSalt[0] ^= 1;
            byte[] otherFirst = EncryptPage(first, 1, rawKey, otherSalt);
            File.WriteAllBytes(source, otherFirst.Concat(EncryptPage(second, 2, rawKey, otherSalt)).ToArray());
            Expect<CryptographicException>(() => decryptor.DecryptDatabase(source, output, key, true), "reject replaced database salt");
            Expect<CryptographicException>(() => decryptor.DecryptPage(otherFirst, 1, key), "reject page salt mismatch");
            File.WriteAllBytes(source, encryptedFirst.Concat(encryptedSecond).ToArray());

            byte[] committed = PlainPage(2, 33);
            WriteWal(source + "-wal", EncryptPage(committed, 2, rawKey, salt),
                EncryptPage(PlainPage(2, 44), 2, rawKey, salt));
            decryptor.DecryptDatabase(source, output, key, true);
            Check(File.ReadAllBytes(output).SequenceEqual(first.Concat(committed)), "apply committed WAL; ignore pending frame");
            decryptor.DecryptDatabase(source, output, key, true);
            Check(File.ReadAllBytes(output).SequenceEqual(first.Concat(committed)), "repeat refresh with WAL");
            File.Delete(source + "-wal");

            decryptor.DecryptFile(source, output, rawKey, true);
            Check(File.ReadAllBytes(output).SequenceEqual(expected), "legacy raw-key API");
            Check(rawKey[0] == 1 && rawKey[31] == 32, "legacy API preserves caller key");
        }
        finally { key.Dispose(); }
        key.Dispose();
        Expect<ObjectDisposedException>(() => decryptor.DecryptPage(encryptedFirst, 1, key), "disposed page key");
        Expect<ObjectDisposedException>(() => decryptor.DecryptDatabase(source, output, key, true), "disposed database key");

        using (var codec = new MemoryCodec(rawKey, salt))
        {
            int pid = Process.GetCurrentProcess().Id;
            var keys = decryptor.DiscoverKeys(pid, new[] { info, info, decryptor.ReadDatabaseInfo(copy) });
            try
            {
                Check(keys.Count == 2, "single scan handles duplicate paths and shared salts");
                Check(keys[source] != keys[copy], "each returned key has independent ownership");
                keys[source].Dispose();
                Check(decryptor.DecryptPage(encryptedFirst, 1, keys[copy]).SequenceEqual(first), "disposing one key preserves another");
            }
            finally { foreach (var discovered in keys.Values) discovered.Dispose(); }
            using (var discovered = decryptor.DiscoverKey(pid, info))
            {
                Check(discovered != null, "single key discovery");
                Check(decryptor.DecryptPage(encryptedFirst, 1, discovered).SequenceEqual(first), "discovered key decrypts");
            }
            string batchOutput = Path.Combine(directory, "batch");
            var results = decryptor.DecryptAvailable(directory, batchOutput, pid);
            Check(results.Count == 2 && results.All(r => r.Status == "Decrypted"), "legacy batch API");
            Check(decryptor.DecryptAvailable(directory, batchOutput, pid).All(r => r.Status == "SkippedExisting"), "legacy batch skip");
        }
        Expect<System.ComponentModel.Win32Exception>(() => decryptor.DiscoverKey(int.MaxValue, info), "invalid process");
        File.WriteAllBytes(source, new byte[5]);
        Expect<InvalidDataException>(() => decryptor.ReadDatabaseInfo(source), "invalid database size");
        string malformed = Path.Combine(directory, "malformed");
        Directory.CreateDirectory(malformed);
        File.WriteAllBytes(Path.Combine(malformed, "chatLogs_1.edb"), new byte[5]);
        var failures = decryptor.DecryptAvailable(malformed, Path.Combine(directory, "failed-output"), int.MaxValue);
        Check(failures.Count == 1 && failures[0].Status == "Failed", "legacy batch reports invalid database without scanning");
    }

    private static byte[] PlainPage(int number, byte marker)
    {
        byte[] page = new byte[PageSize];
        for (int i = 0; i < PageSize - 80; i++) page[i] = marker;
        if (number == 1)
        {
            Buffer.BlockCopy(Encoding.ASCII.GetBytes("SQLite format 3\0"), 0, page, 0, 16);
            page[16] = 16;
            page[17] = 0;
            page[20] = 80;
        }
        return page;
    }

    private static byte[] EncryptPage(byte[] plain, int number, byte[] key, byte[] salt)
    {
        byte[] page = new byte[PageSize];
        int start = number == 1 ? 16 : 0;
        if (start != 0) Buffer.BlockCopy(salt, 0, page, 0, 16);
        byte[] iv = Enumerable.Range(101, 16).Select(i => (byte)i).ToArray();
        using (var aes = Aes.Create())
        {
            aes.Key = key;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using (var encryptor = aes.CreateEncryptor())
            {
                byte[] cipher = encryptor.TransformFinalBlock(plain, start, PageSize - 80 - start);
                Buffer.BlockCopy(cipher, 0, page, start, cipher.Length);
            }
        }
        Buffer.BlockCopy(iv, 0, page, PageSize - 80, 16);
        // Independent standard PBKDF2 implementation, rather than production's MacKey helper.
        using (var pbkdf = new Rfc2898DeriveBytes(key, salt.Select(b => (byte)(b ^ 0x3a)).ToArray(), 2, HashAlgorithmName.SHA512))
        using (var hmac = new HMACSHA512(pbkdf.GetBytes(32)))
        {
            byte[] authenticated = page.Skip(start).Take(PageSize - 64 - start)
                .Concat(BitConverter.GetBytes((uint)number)).ToArray();
            Buffer.BlockCopy(hmac.ComputeHash(authenticated), 0, page, PageSize - 64, 64);
        }
        return page;
    }

    private static void WriteWal(string path, byte[] committed, byte[] pending)
    {
        byte[] header = new byte[32];
        WriteBigEndian(header, 0, 0x377f0682);
        WriteBigEndian(header, 8, PageSize);
        WriteBigEndian(header, 16, 123);
        WriteBigEndian(header, 20, 456);
        using (var stream = File.Create(path))
        {
            stream.Write(header, 0, header.Length);
            foreach (var frame in new[] { committed, pending })
            {
                byte[] frameHeader = new byte[24];
                WriteBigEndian(frameHeader, 0, 2);
                WriteBigEndian(frameHeader, 4, frame == committed ? 2U : 0U);
                Buffer.BlockCopy(header, 16, frameHeader, 8, 8);
                stream.Write(frameHeader, 0, frameHeader.Length);
                stream.Write(frame, 0, frame.Length);
            }
        }
    }

    private static void WriteBigEndian(byte[] bytes, int offset, uint value)
    {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (24 - i * 8));
    }

    private static void Check(bool condition, string name)
    {
        if (!condition) throw new Exception("FAIL: " + name);
        passed++;
    }

    private static void Expect<T>(Action action, string name) where T : Exception
    {
        try { action(); }
        catch (T) { passed++; return; }
        throw new Exception("FAIL: " + name + " did not throw " + typeof(T).Name);
    }

    // Mimic the existing codec layout in our own process; never inspects a real KakaoTalk process.
    private sealed class MemoryCodec : IDisposable
    {
        private readonly IntPtr context = Marshal.AllocHGlobal(128);
        private readonly IntPtr cipher = Marshal.AllocHGlobal(32);
        private readonly IntPtr keyMemory = Marshal.AllocHGlobal(32);
        private readonly IntPtr saltMemory = Marshal.AllocHGlobal(16);

        public MemoryCodec(byte[] key, byte[] salt)
        {
            Marshal.Copy(new byte[128], 0, context, 128);
            Marshal.Copy(new byte[32], 0, cipher, 32);
            Marshal.Copy(key, 0, keyMemory, 32);
            Marshal.Copy(salt, 0, saltMemory, 16);
            int[] signature = { 256000, 2, 16, 32, 16, 16, PageSize };
            for (int i = 0; i < signature.Length; i++) Marshal.WriteInt32(context, 4 + 4 * i, signature[i]);
            Marshal.WriteInt32(context, 36, 80);
            Marshal.WriteInt32(context, 40, 64);
            Marshal.WriteIntPtr(context, 64, saltMemory);
            Marshal.WriteIntPtr(context, 96, cipher);
            Marshal.WriteIntPtr(cipher, 8, keyMemory);
        }

        public void Dispose()
        {
            Marshal.Copy(new byte[32], 0, keyMemory, 32);
            Marshal.FreeHGlobal(context);
            Marshal.FreeHGlobal(cipher);
            Marshal.FreeHGlobal(keyMemory);
            Marshal.FreeHGlobal(saltMemory);
        }
    }
}
