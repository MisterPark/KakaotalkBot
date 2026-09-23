using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace KakaotalkBot
{
    // Windows에 포함된 SQLite를 사용합니다. 조회 대상은 봇이 만든 복호화 사본뿐입니다.
    internal sealed class ChatSqlite : IDisposable
    {
        private IntPtr database;

        public ChatSqlite(string path, bool readOnly = true)
        {
            string name = readOnly ? new Uri(Path.GetFullPath(path)).AbsoluteUri + "?immutable=1" : path;
            int result = sqlite3_open_v2(Utf8(name), out database, readOnly ? 0x41 : 6, IntPtr.Zero);
            if (result != 0)
            {
                string message = Error();
                Dispose();
                throw new IOException("채팅 SQLite 열기 실패: " + message);
            }
        }

        public List<object[]> Query(string sql, params object[] arguments)
        {
            IntPtr statement;
            int result = sqlite3_prepare_v2(database, Utf8(sql), -1, out statement, IntPtr.Zero);
            if (result != 0) throw new IOException("채팅 SQLite 조회 준비 실패: " + Error());
            try
            {
                for (int i = 0; i < arguments.Length; i++)
                {
                    if (arguments[i] == null) result = sqlite3_bind_null(statement, i + 1);
                    else if (arguments[i] is string)
                    {
                        byte[] text = Utf8((string)arguments[i]);
                        result = sqlite3_bind_text(statement, i + 1, text, text.Length - 1, new IntPtr(-1));
                    }
                    else result = sqlite3_bind_int64(statement, i + 1, Convert.ToInt64(arguments[i]));
                    if (result != 0) throw new IOException("채팅 SQLite 매개변수 오류: " + Error());
                }
                var rows = new List<object[]>();
                while ((result = sqlite3_step(statement)) == 100)
                {
                    var row = new object[sqlite3_column_count(statement)];
                    for (int i = 0; i < row.Length; i++)
                    {
                        int type = sqlite3_column_type(statement, i);
                        row[i] = type == 5 ? null : type == 1 ? (object)sqlite3_column_int64(statement, i)
                            : ReadUtf8(sqlite3_column_text(statement, i), sqlite3_column_bytes(statement, i));
                    }
                    rows.Add(row);
                }
                if (result != 101) throw new IOException("채팅 SQLite 조회 실패: " + Error());
                return rows;
            }
            finally { sqlite3_finalize(statement); }
        }

        public long Scalar(string sql, params object[] arguments)
        {
            var rows = Query(sql, arguments);
            return rows.Count == 0 || rows[0][0] == null ? 0 : Convert.ToInt64(rows[0][0]);
        }

        private string Error()
        {
            IntPtr pointer = database == IntPtr.Zero ? IntPtr.Zero : sqlite3_errmsg(database);
            if (pointer == IntPtr.Zero) return "SQLite 초기화 오류";
            int length = 0;
            while (Marshal.ReadByte(pointer, length) != 0) length++;
            return ReadUtf8(pointer, length);
        }

        private static byte[] Utf8(string value) { return Encoding.UTF8.GetBytes(value + "\0"); }
        private static string ReadUtf8(IntPtr pointer, int length)
        {
            byte[] bytes = new byte[length];
            if (length > 0) Marshal.Copy(pointer, bytes, 0, length);
            return Encoding.UTF8.GetString(bytes);
        }

        public void Dispose()
        {
            if (database != IntPtr.Zero) { sqlite3_close(database); database = IntPtr.Zero; }
        }

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_open_v2(byte[] path, out IntPtr db, int flags, IntPtr vfs);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_close(IntPtr db);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr sqlite3_errmsg(IntPtr db);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int bytes, out IntPtr statement, IntPtr tail);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_step(IntPtr statement);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_finalize(IntPtr statement);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_bind_null(IntPtr statement, int index);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_bind_int64(IntPtr statement, int index, long value);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_bind_text(IntPtr statement, int index, byte[] value, int bytes, IntPtr destructor);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_column_count(IntPtr statement);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_column_type(IntPtr statement, int column);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern long sqlite3_column_int64(IntPtr statement, int column);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr sqlite3_column_text(IntPtr statement, int column);
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int sqlite3_column_bytes(IntPtr statement, int column);
    }
}
