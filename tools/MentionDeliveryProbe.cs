using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using KakaotalkBot;

internal static class MentionDeliveryProbe
{
    private const string Room = "보룸봇 테스트방";
    private const string ChatFile = "chatLogs_18483459527844956.edb";
    private const string Message = "@채팅봇 멘션 ID 검증";
    private const long BeforeId = 6593948531989662102L, AfterId = 8172472438151128507L;
    private static readonly StringBuilder Report = new StringBuilder();
    private static string output;

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try
        {
            if (args.Length < 1 || args.Length > 2) throw new ArgumentException("baseline / prepare / restore / verify 기준logId");
            output = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Diagnostics", DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(output);
            Line("수집 시각: " + DateTimeOffset.Now.ToString("O"));
            if (args[0] == "prepare" || args[0] == "restore")
            {
                bool restore = args[0] == "restore";
                // 승인된 시험 입력과 멘션 하나에만 적용한다. 이 도구는 키 입력이나 전송을 하지 않는다.
                var snapshot = KakaoInputMentions.ReplaceUserId(Room, 0, restore ? AfterId : BeforeId, restore ? BeforeId : AfterId,
                    "\uFFFC 멘션 ID 검증\r", "@채팅봇");
                Line("표시=" + snapshot.Mentions[0].DisplayText + ", 대상 ID=" + snapshot.Mentions[0].UserId);
                Save(); return 0;
            }
            if (args[0] != "baseline" && args[0] != "verify") throw new ArgumentException("지원하지 않는 명령입니다.");
            long baseline = 0;
            if (args[0] == "verify" && (args.Length != 2 || !long.TryParse(args[1], NumberStyles.None, CultureInfo.InvariantCulture, out baseline)))
                throw new ArgumentException("검증하려면 전송 전 기준 logId가 필요합니다.");
            var source = KakaoTalkDecryptor.FindChatDataDirectories().Select(path => Path.Combine(path, ChatFile)).Where(File.Exists).ToArray();
            if (source.Length != 1) throw new InvalidOperationException("대상 채팅방의 원본 DB를 유일하게 확인하지 못했습니다.");
            var input = KakaoInputMentions.Read(Room);
            var decryptor = new KakaoTalkDecryptor();
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(45)))
            using (var key = decryptor.DiscoverKey(input.ProcessId, decryptor.ReadDatabaseInfo(source[0]), cancellation.Token))
            {
                if (key == null) throw new InvalidOperationException("대상 DB의 재사용 키를 찾지 못했습니다.");
                var snapshot = new ChatDatabaseSnapshot(source[0], Path.Combine(output, "chat.sqlite"), decryptor, key);
                for (int attempt = 0; attempt < 20; attempt++)
                {
                    snapshot.Refresh(cancellation.Token);
                    using (var db = new ChatSqlite(snapshot.OutputPath))
                    {
                        if (args[0] == "baseline")
                        {
                            Line("기준 logId=" + db.Scalar("SELECT MAX(logId) FROM chatLogs"));
                            Line("기존 동일 메시지 수=" + db.Scalar("SELECT COUNT(*) FROM chatLogs WHERE message = ?", Message));
                            Save(); return 0;
                        }
                        var rows = db.Query("SELECT logId, authorId, type, sendAt, message, attachement FROM chatLogs WHERE logId > ? AND message = ? ORDER BY logId", baseline, Message);
                        if (rows.Count > 1) throw new InvalidOperationException("기준 이후 동일 메시지가 여러 개 있어 단일 전송을 확인할 수 없습니다.");
                        if (rows.Count == 1)
                        {
                            var row = rows[0];
                            Line("logId=" + row[0] + ", authorId=" + row[1] + ", type=" + row[2] + ", sendAt=" + row[3]);
                            Line("message=" + row[4]); Line("attachement=" + row[5]);
                            var data = new JavaScriptSerializer().DeserializeObject(Convert.ToString(row[5])) as Dictionary<string, object>;
                            object mentions;
                            var ids = new List<long>();
                            if (data != null && data.TryGetValue("mentions", out mentions))
                                foreach (object mention in (IEnumerable)mentions)
                                {
                                    var item = mention as Dictionary<string, object>; object id;
                                    if (item != null && item.TryGetValue("user_id", out id)) ids.Add(Convert.ToInt64(id, CultureInfo.InvariantCulture));
                                }
                            bool success = ids.Count == 1 && ids[0] == AfterId;
                            Line(success ? "PASS: 전송 후 DB의 멘션 대상이 보룸봇 ID와 일치합니다." : "FAIL: 전송 후 DB의 멘션 대상이 예상과 다릅니다.");
                            Save(); return success ? 0 : 2;
                        }
                    }
                    Thread.Sleep(1000);
                }
            }
            Line("미확인: 기준 이후 승인된 시험 메시지가 아직 DB에 없습니다. 자동 재전송하지 않습니다.");
            Save(); return 3;
        }
        catch (Exception error) { Line("실패: " + error.Message); if (output != null) Save(); return 1; }
    }
    private static void Line(string text) { Report.Append(text).Append("\r\n"); Console.WriteLine(text); }
    private static void Save() { string path = Path.Combine(output, "report.txt"); File.WriteAllText(path, Report.ToString(), new UTF8Encoding(true)); Console.WriteLine("보고서: " + path); }
}
