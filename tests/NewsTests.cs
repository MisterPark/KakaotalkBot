using System;
using System.Net.Http;
using System.Threading.Tasks;
using KakaotalkBot;
internal static class NewsTests
{
    static void Check(bool ok, string name) { if (!ok) throw new Exception(name); }
    static string Html()
    {
        string html = "<html><body>";
        for (int i = 0; i < 8; i++) html += "<a class='sa_text_title other' href='https://example.com/" + i + "'><strong>기사 &amp; " + i + "</strong></a>";
        return html + "</body></html>";
    }
    static int Main()
    {
        try
        {
            Check(News.PoliticsTop6.Contains("불러오지 못"), "초기 안내");
            News.UpdatePoliticsTop6(() => Task.FromResult(Html())).GetAwaiter().GetResult();
            var original = News.PoliticsTop6;
            Check(original.Contains("6. 기사 & 5") && !original.Contains("7. 기사"), "6개와 HTML 디코딩");
            var failed = new TaskCompletionSource<string>(); failed.SetException(new HttpRequestException("시험 연결 실패"));
            News.UpdatePoliticsTop6(() => failed.Task).GetAwaiter().GetResult();
            Check(News.PoliticsTop6 == original && News.LastRefreshError.Contains("통신 실패"), "HTTP 실패 시 보존");
            var canceled = new TaskCompletionSource<string>(); canceled.SetCanceled();
            News.UpdatePoliticsTop6(() => canceled.Task).GetAwaiter().GetResult();
            Check(News.PoliticsTop6 == original && News.LastRefreshError.Contains("시간 초과"), "취소 처리");
            News.UpdatePoliticsTop6(() => Task.FromResult("<html>차단 안내</html>")).GetAwaiter().GetResult();
            Check(News.PoliticsTop6 == original && News.LastRefreshError != null, "빈 파싱 보존");
            var pending = new TaskCompletionSource<string>();
            var running = News.UpdatePoliticsTop6(() => pending.Task);
            int duplicate = 0;
            News.UpdatePoliticsTop6(() => { duplicate++; return Task.FromResult(Html()); }).GetAwaiter().GetResult();
            Check(duplicate == 0, "중복 요청 차단");
            pending.SetResult(Html()); running.GetAwaiter().GetResult();
            Check(News.LastRefreshError == null && News.PoliticsTop6 == original, "다음 갱신 복구");
            Console.WriteLine("PASS: HTTP failure, cancellation, invalid HTML, cache preservation, six articles, concurrent refresh and recovery (no network)");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
