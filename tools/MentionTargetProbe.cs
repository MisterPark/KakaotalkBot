using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using KakaotalkBot;

internal static class MentionTargetProbe
{
    [STAThread]
    private static int Main(string[] args)
    {
        var report = new StringBuilder();
        try
        {
            if (args.Length != 1 && args.Length != 5) throw new ArgumentException("사용법: MentionTargetProbe.exe 방이름 [--round-trip 위치 현재ID 시험ID]");
            int position = 0; long expected = 0, replacement = 0;
            if (args.Length == 5 && (args[1] != "--round-trip" || !int.TryParse(args[2], out position) || position < 0 ||
                !long.TryParse(args[3], NumberStyles.None, CultureInfo.InvariantCulture, out expected) || expected <= 0 ||
                !long.TryParse(args[4], NumberStyles.None, CultureInfo.InvariantCulture, out replacement) || replacement <= 0 || expected == replacement))
                throw new ArgumentException("시험 인자를 확인하세요.");
            report.AppendLine("수집 시각: " + DateTimeOffset.Now.ToString("O"));
            for (int iteration = 0; iteration < 2; iteration++)
            {
                var clock = Stopwatch.StartNew();
                var snapshot = KakaoInputMentions.Read(args[0]);
                report.AppendLine("조회 " + (iteration + 1) + ": " + clock.ElapsedMilliseconds + " ms, 멘션 " + snapshot.Mentions.Count + "개");
                foreach (var mention in snapshot.Mentions) report.AppendLine("  위치=" + mention.Position + ", ID=" + mention.UserId + ", 표시=" + mention.DisplayText);
            }
            if (args.Length == 5) report.AppendLine(KakaoInputMentions.VerifyUserIdWriteAndRestore(args[0], position, expected, replacement));
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Diagnostics", DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "report.txt");
            File.WriteAllText(path, report.ToString(), new UTF8Encoding(true));
            Console.Write(report); Console.WriteLine("보고서: " + path);
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(report.ToString() + "실패: " + error.Message); return 1; }
    }
}
