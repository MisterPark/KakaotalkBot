using System;
using System.Diagnostics;
using KakaotalkBot;

internal static class MentionPreparationProbe
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length != 3) throw new ArgumentException("방 이름, 대상 사용자 ID, 후보 검색용 닉네임을 지정하세요.");
            long id;
            if (!long.TryParse(args[1], out id) || id <= 0) throw new ArgumentException("올바른 사용자 ID가 필요합니다.");
            var sender = new KakaoMentionSender(args[0]);
            var clock = Stopwatch.StartNew();
            var prepared = sender.Prepare(id, args[2], "자동 준비 검증\n둘째 줄");
            try
            {
                Console.WriteLine("준비 완료: 대상=" + prepared.Mentions[0].UserId + ", 원본 표시=" + prepared.Mentions[0].DisplayText + ", 내용=" + prepared.Text.Replace("\r", "\\r"));
                Console.WriteLine("준비 소요: " + clock.ElapsedMilliseconds + "ms");
            }
            finally { sender.DiscardPrepared(prepared); }
            if (!KakaoInputMentions.Read(args[0]).IsEmpty) throw new InvalidOperationException("시험 입력 정리 상태를 확인하지 못했습니다.");
            Console.WriteLine("시험 입력 정리 완료. 전송하지 않았습니다.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error.Message); return 1; }
    }
}
