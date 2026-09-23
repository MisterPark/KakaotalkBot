using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using KakaotalkBot;

internal static class LocoSenderBridgeTests
{
    private static int checks;
    private static void Check(bool condition, string name)
    {
        if (!condition) throw new Exception(name);
        checks++;
    }

    private static async Task Run(string node, string bridge)
    {
        using (var sender = new LocoMentionSender(node, bridge))
        {
            Check((int)(await sender.GetInfoAsync())["protocol"] == 1, "프로토콜 확인");
            Check((bool)(await sender.GetInfoAsync())["registrationSupported"] == false, "신규 기기 등록 미지원 표시");
            try { await sender.RequestPasscodeAsync("test@example.invalid", "가짜비밀번호"); throw new Exception("미지원 등록 요청 허용"); }
            catch (LocoSenderException error) { Check(error.Code == "REGISTRATION_UNSUPPORTED", "등록 미지원 오류 전달"); }
            var parts = new[] { LocoMessagePart.Mention(long.MaxValue, "보룸봇😀"), LocoMessagePart.Text(" 한글 본문") };
            var preview = await sender.PreviewAsync(9007199254740993L, parts);
            Check((string)preview["chatId"] == "9007199254740993", "방 ID 정밀도");
            Check((string)preview["mentions"][0]["userId"] == long.MaxValue.ToString(), "대상 ID 정밀도");
            Check((string)preview["text"] == "@보룸봇😀 한글 본문", "UTF-8 본문");
            var parallel = await Task.WhenAll(Enumerable.Range(0, 4).Select(i => sender.PreviewAsync(i + 1, parts)));
            Check(parallel.Select(value => (string)value["chatId"]).SequenceEqual(new[] { "1", "2", "3", "4" }), "동시 요청 직렬 처리");
            try { await sender.PrepareAsync(1, parts); throw new Exception("미로그인 전송 준비 허용"); }
            catch (LocoSenderException error) { Check(error.Code == "NOT_CONNECTED", "미로그인 거절"); }
        }

        // 전송 직후 프로세스가 종료되면 성공/실패를 추정하지 않고 불명 상태를 돌려줍니다.
        string directory = Path.Combine(Path.GetTempPath(), "LocoBridgeTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string fake = Path.Combine(directory, "exit.js");
            File.WriteAllText(fake, "process.stdin.once('data', () => process.exit(1));\r\n");
            using (var sender = new LocoMentionSender(node, fake))
            {
                try { await sender.SendPreparedAsync("test-ticket"); throw new Exception("미확인 전송 성공 처리"); }
                catch (LocoSenderException error) { Check(error.Code == "SEND_UNKNOWN", "전송 결과 불명"); }
            }
        }
        finally { Directory.Delete(directory, true); }
    }

    private static int Main(string[] args)
    {
        try { Run(args[0], args[1]).GetAwaiter().GetResult(); Console.WriteLine("C# LOCO bridge: " + checks + " checks passed"); return 0; }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
