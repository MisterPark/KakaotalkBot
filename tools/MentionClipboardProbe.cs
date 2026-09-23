using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace KakaoMentionProbe
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Contains("--self-test")) return ProbeAnalysis.SelfTest();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ProbeForm());
            return 0;
        }
    }

    internal sealed class ProbeForm : Form
    {
        private readonly TextBox expectedId = new TextBox();
        private readonly TextBox report = new TextBox();
        private readonly string outputDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Diagnostics",
            DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
        private ClipboardSample plain, mention;

        public ProbeForm()
        {
            Text = "카카오톡 멘션 클립보드 진단";
            ClientSize = new Size(850, 580);
            MinimumSize = new Size(700, 500);
            Font = new Font("맑은 고딕", 10);
            var guide = new Label { AutoSize = false, Text =
                "카카오톡 입력창에서 내용을 선택해 Ctrl+C로 복사한 뒤 아래 버튼을 누르세요.\r\n" +
                "일반 @이름과 후보에서 선택한 실제 멘션을 각각 수집하면 비교할 수 있습니다. 전송은 필요 없습니다.",
                Location = new Point(12, 12), Size = new Size(825, 54), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            var idLabel = new Label { Text = "비교할 user_id (선택)", AutoSize = true, Location = new Point(12, 76) };
            expectedId.SetBounds(190, 72, 245, 28);
            var plainButton = new Button { Text = "1. 일반 텍스트 수집", Location = new Point(12, 110), Size = new Size(180, 34) };
            var mentionButton = new Button { Text = "2. 선택된 멘션 수집", Location = new Point(202, 110), Size = new Size(180, 34) };
            var compareButton = new Button { Text = "ID 비교 다시 실행", Location = new Point(392, 110), Size = new Size(170, 34) };
            var folderButton = new Button { Text = "결과 폴더 열기", Location = new Point(572, 110), Size = new Size(170, 34) };
            plainButton.Click += (sender, e) => CaptureSample(false);
            mentionButton.Click += (sender, e) => CaptureSample(true);
            compareButton.Click += (sender, e) => UpdateReport();
            folderButton.Click += (sender, e) =>
            {
                if (Directory.Exists(outputDirectory)) Process.Start(new ProcessStartInfo(outputDirectory) { UseShellExecute = true });
            };
            report.Multiline = true;
            report.ReadOnly = true;
            report.ScrollBars = ScrollBars.Both;
            report.WordWrap = false;
            report.SetBounds(12, 157, 825, 410);
            report.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            report.Text = "수집 대기 중. 카카오톡이 복사한 클립보드만 읽습니다.\r\n" +
                "RTF/HTML/사용자 정의 형식의 원본 바이트와 보고서는 exe 하위 Diagnostics에 저장됩니다.\r\n" +
                "이 도구는 붙여넣기, 키 입력, 메시지 전송을 수행하지 않습니다.";
            Controls.AddRange(new Control[] { guide, idLabel, expectedId, plainButton, mentionButton, compareButton, folderButton, report });
        }

        private void CaptureSample(bool selectedMention)
        {
            try
            {
                var sample = ClipboardReader.ReadKakao();
                string stage = (selectedMention ? "mention-" : "plain-") + Guid.NewGuid().ToString("N").Substring(0, 8);
                string directory = Path.Combine(outputDirectory, stage);
                Directory.CreateDirectory(directory);
                foreach (var format in sample.Formats.Where(format => format.Data != null))
                    File.WriteAllBytes(Path.Combine(directory, format.Id + ".bin"), format.Data);
                File.WriteAllText(Path.Combine(directory, "formats.txt"), ProbeAnalysis.Describe(sample, null), new UTF8Encoding(true));
                if (selectedMention) mention = sample; else plain = sample;
                UpdateReport();
            }
            catch (Exception ex) { report.Text = "수집 실패: " + ex.Message; }
        }

        private void UpdateReport()
        {
            long id;
            long? target = null;
            if (!string.IsNullOrWhiteSpace(expectedId.Text))
            {
                if (!long.TryParse(expectedId.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out id) || id <= 0)
                { report.Text = "user_id는 양의 정수로 입력하거나 비워 두세요."; return; }
                target = id;
            }
            var text = new StringBuilder();
            text.AppendLine("저장 위치: " + outputDirectory);
            text.AppendLine("[일반 텍스트]").AppendLine(plain == null ? "미수집" : ProbeAnalysis.Describe(plain, target));
            text.AppendLine("[선택된 멘션]").AppendLine(mention == null ? "미수집" : ProbeAnalysis.Describe(mention, target));
            if (plain != null && mention != null)
            {
                text.AppendLine("[비교]");
                if (plain.Sequence == mention.Sequence) text.AppendLine("동일한 복사본입니다. 다른 입력 상태에서 다시 복사해 주세요.");
                text.AppendLine("표시 문자열 일치: " + (plain.UnicodeText == null || mention.UnicodeText == null ?
                    "유니코드 텍스트가 없어 미확인" : (plain.UnicodeText == mention.UnicodeText).ToString()));
                foreach (var format in mention.Formats)
                {
                    var before = plain.Formats.FirstOrDefault(value => value.Id == format.Id);
                    text.AppendLine(format.Name + ": " + (before == null ? "멘션에만 존재" :
                        before.Data == null || format.Data == null ? "원본 비교 불가" :
                        before.Data.SequenceEqual(format.Data) ? "동일" : "변경됨"));
                }
                foreach (var format in plain.Formats.Where(value => !mention.Formats.Any(after => after.Id == value.Id)))
                    text.AppendLine(format.Name + ": 일반 텍스트에만 존재");
            }
            text.AppendLine("ID 바이트가 발견되어도 대상 지정 성공을 뜻하지 않습니다. 멘션 형식 해석과 별도 검증이 필요합니다.");
            report.Text = text.ToString();
            if (plain != null || mention != null)
            {
                Directory.CreateDirectory(outputDirectory);
                File.WriteAllText(Path.Combine(outputDirectory, "report.txt"), report.Text, new UTF8Encoding(true));
            }
        }
    }

    internal sealed class ClipboardSample
    {
        public uint Sequence;
        public readonly List<ClipboardFormat> Formats = new List<ClipboardFormat>();
        public string UnicodeText
        {
            get
            {
                var format = Formats.FirstOrDefault(value => value.Id == 13 && value.Data != null);
                return format == null ? null : Encoding.Unicode.GetString(format.Data).TrimEnd('\0');
            }
        }
    }

    internal sealed class ClipboardFormat
    {
        public uint Id;
        public string Name, Note;
        public byte[] Data;
    }

    internal static class ClipboardReader
    {
        // 형식별 4MiB, 총 16MiB까지만 읽고 객체 역직렬화는 하지 않습니다.
        private const ulong MaxFormatBytes = 4 * 1024 * 1024;
        public static ClipboardSample ReadKakao()
        {
            if (!OpenClipboard(IntPtr.Zero)) throw new IOException("클립보드를 다른 앱이 사용 중입니다. 다시 시도하세요.");
            try
            {
                uint processId;
                GetWindowThreadProcessId(GetClipboardOwner(), out processId);
                if (processId == 0) throw new IOException("복사한 앱을 확인할 수 없습니다. 카카오톡에서 다시 Ctrl+C를 누르세요.");
                using (var process = Process.GetProcessById((int)processId))
                    if (!string.Equals(process.ProcessName, "KakaoTalk", StringComparison.OrdinalIgnoreCase))
                        throw new IOException("카카오톡에서 복사한 데이터가 아닙니다. 카카오톡 입력창의 멘션을 다시 복사하세요.");
                var sample = new ClipboardSample { Sequence = GetClipboardSequenceNumber() };
                var ids = new List<uint>();
                uint current = 0;
                while ((current = EnumClipboardFormats(current)) != 0)
                {
                    if (ids.Count >= 128 || ids.Contains(current)) throw new IOException("클립보드 형식 목록이 비정상적입니다.");
                    ids.Add(current);
                }
                ulong total = 0;
                foreach (uint id in ids)
                {
                    var name = new StringBuilder(256);
                    GetClipboardFormatName(id, name, name.Capacity);
                    var format = new ClipboardFormat { Id = id, Name = name.Length == 0 ? StandardName(id) : name.ToString() };
                    sample.Formats.Add(format);
                    // 텍스트·RTF·등록된 전용 형식만 수집합니다. 이미지/GDI/파일 목록은 건드리지 않습니다.
                    if (!(id == 1 || id == 7 || id == 13 || id == 16 || id >= 0xC000))
                    { format.Note = "원본 읽기 제외"; continue; }
                    IntPtr handle = GetClipboardData(id);
                    if (handle == IntPtr.Zero) { format.Note = "원본 읽기 실패"; continue; }
                    ulong size = GlobalSize(handle).ToUInt64();
                    if (size == 0 || size > MaxFormatBytes || total + size > 16 * 1024 * 1024)
                    { format.Note = "메모리 형식 미확인 또는 크기 제한 초과"; continue; }
                    IntPtr pointer = GlobalLock(handle);
                    if (pointer == IntPtr.Zero) { format.Note = "메모리 잠금 실패"; continue; }
                    try
                    {
                        format.Data = new byte[(int)size];
                        Marshal.Copy(pointer, format.Data, 0, format.Data.Length);
                        total += size;
                    }
                    finally { GlobalUnlock(handle); }
                }
                return sample;
            }
            finally { CloseClipboard(); }
        }

        private static string StandardName(uint id)
        { return id == 1 ? "CF_TEXT" : id == 7 ? "CF_OEMTEXT" : id == 13 ? "CF_UNICODETEXT" : id == 16 ? "CF_LOCALE" : "Format_" + id; }

        [DllImport("user32.dll", SetLastError = true)] private static extern bool OpenClipboard(IntPtr window);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")] private static extern uint EnumClipboardFormats(uint format);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClipboardFormatName(uint format, StringBuilder name, int count);
        [DllImport("user32.dll")] private static extern IntPtr GetClipboardData(uint format);
        [DllImport("kernel32.dll")] private static extern UIntPtr GlobalSize(IntPtr handle);
        [DllImport("kernel32.dll")] private static extern IntPtr GlobalLock(IntPtr handle);
        [DllImport("kernel32.dll")] private static extern bool GlobalUnlock(IntPtr handle);
    }

    internal static class ProbeAnalysis
    {
        public static string Describe(ClipboardSample sample, long? target)
        {
            var text = new StringBuilder();
            text.AppendLine("클립보드 순번: " + sample.Sequence);
            foreach (var format in sample.Formats)
            {
                text.AppendLine(format.Id + " " + format.Name + " / " + (format.Data == null ? format.Note : format.Data.Length + "바이트"));
                if (format.Data == null) continue;
                using (var sha = SHA256.Create()) text.AppendLine("  SHA256: " + BitConverter.ToString(sha.ComputeHash(format.Data)).Replace("-", ""));
                if (target.HasValue)
                {
                    string id = target.Value.ToString(CultureInfo.InvariantCulture);
                    byte[] littleEndian = BitConverter.GetBytes(target.Value);
                    byte[] bigEndian = littleEndian.Reverse().ToArray();
                    var matches = new List<string>();
                    if (Contains(format.Data, Encoding.ASCII.GetBytes(id))) matches.Add("십진 문자열");
                    if (Contains(format.Data, Encoding.Unicode.GetBytes(id))) matches.Add("UTF-16 문자열");
                    if (Contains(format.Data, littleEndian)) matches.Add("64비트 LE");
                    if (Contains(format.Data, bigEndian)) matches.Add("64비트 BE");
                    text.AppendLine("  대상 ID 바이트 후보: " + (matches.Count == 0 ? "없음" : string.Join(", ", matches)));
                }
                bool textual = format.Id == 13 || format.Id == 1 || format.Id == 7 ||
                    format.Name.IndexOf("Rich Text", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    format.Name.IndexOf("HTML", StringComparison.OrdinalIgnoreCase) >= 0;
                if (textual)
                {
                    Encoding encoding = format.Id == 13 ? Encoding.Unicode : format.Id == 1 ? Encoding.Default :
                        format.Id == 7 ? Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage) : Encoding.UTF8;
                    string value = encoding.GetString(format.Data, 0, Math.Min(format.Data.Length, 16384)).TrimEnd('\0');
                    text.AppendLine("  데이터 미리보기: " + value.Replace("\r", "\\r").Replace("\n", "\\n"));
                }
            }
            return text.ToString();
        }

        internal static bool Contains(byte[] data, byte[] pattern)
        {
            if (pattern.Length == 0) return false;
            for (int i = 0; i <= data.Length - pattern.Length; i++)
            {
                int j = 0;
                while (j < pattern.Length && data[i + j] == pattern[j]) j++;
                if (j == pattern.Length) return true;
            }
            return false;
        }

        internal static int SelfTest()
        {
            // 실제 클립보드나 카카오톡을 조작하지 않는 분석기 검사입니다.
            const long id = 9007199254740993L;
            var sample = new ClipboardSample();
            sample.Formats.Add(new ClipboardFormat { Id = 13, Name = "CF_UNICODETEXT", Data = Encoding.Unicode.GetBytes("@검증\0") });
            sample.Formats.Add(new ClipboardFormat { Id = 50000, Name = "TestPayload", Data = BitConverter.GetBytes(id) });
            string report = Describe(sample, id);
            bool success = sample.UnicodeText == "@검증" && report.Contains("64비트 LE") &&
                !Describe(sample, id + 1).Contains("64비트 LE") &&
                Contains(new byte[] { 0, 1, 2, 3 }, new byte[] { 2, 3 }) &&
                !Contains(new byte[] { 1 }, new byte[] { 1, 2 }) &&
                !Contains(new byte[0], new byte[0]);
            Console.WriteLine(success ? "PASS: clipboard payload analysis (no clipboard access)" : "FAIL: clipboard payload analysis");
            return success ? 0 : 1;
        }
    }
}
