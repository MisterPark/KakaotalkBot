using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace KakaotalkBot
{
    public static class News
    {
        private static volatile List<Article> articles = new List<Article>();
        private static readonly HttpClient httpClient = CreateClient();
        private static readonly SemaphoreSlim updateGate = new SemaphoreSlim(1, 1);
        private static volatile string refreshError;
        public static string LastRefreshError { get { return refreshError; } }

        private static HttpClient CreateClient()
        {
            var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
            return client;
        }

        internal static void BeginUpdate()
        {
            // 내부에서 예상 가능한 실패를 처리하고, 진행 중이면 중복 갱신하지 않습니다.
            Task.Run(() => UpdatePoliticsTop6(() => httpClient.GetStringAsync("https://news.naver.com/breakingnews/section/100/269")));
        }

        public static string PoliticsTop6
        {
            get
            {
                var snapshot = articles;
                var sb = new StringBuilder();
                sb.AppendLine("[정치뉴스 TOP6]");
                sb.Append("\r\n");
                if (snapshot.Count == 0)
                    sb.Append("현재 뉴스를 불러오지 못했습니다. 잠시 후 다시 확인해 주세요.");
                for (int i = 0; i < snapshot.Count; i++)
                {
                    sb.AppendLine((i + 1) + ". " + snapshot[i].Headline);
                    sb.AppendLine(snapshot[i].Link);
                }
                return sb.ToString();
            }
        }

        public static void Update()
        {
            UpdatePoliticsTop6(() => httpClient.GetStringAsync("https://news.naver.com/breakingnews/section/100/269")).GetAwaiter().GetResult();
        }

        internal static async Task UpdatePoliticsTop6(Func<Task<string>> fetch)
        {
            if (!await updateGate.WaitAsync(0).ConfigureAwait(false)) return;
            try
            {
                string html = await fetch().ConfigureAwait(false);
                var next = ParseArticles(html);
                // 검증을 마친 새 목록을 한 번에 교체해 실패 시 기존 뉴스를 보존합니다.
                articles = next;
                refreshError = null;
            }
            catch (HttpRequestException)
            {
                refreshError = "뉴스 조회 통신 실패 · 기존 뉴스 유지, 다음 주기에 재시도";
            }
            catch (OperationCanceledException)
            {
                refreshError = "뉴스 조회 시간 초과 또는 취소 · 기존 뉴스 유지, 다음 주기에 재시도";
            }
            catch (InvalidDataException)
            {
                refreshError = "뉴스 항목을 확인하지 못했습니다 · 기존 뉴스 유지, 다음 주기에 재시도";
            }
            finally { updateGate.Release(); }
        }

        internal static List<Article> ParseArticles(string html)
        {
            if (string.IsNullOrWhiteSpace(html)) throw new InvalidDataException("빈 뉴스 응답");
            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            // 화면의 div 위치 대신 기사 제목의 클래스로 찾습니다.
            var nodes = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', normalize-space(@class), ' '), ' sa_text_title ')]");
            var next = new List<Article>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            if (nodes != null)
            {
                foreach (var node in nodes)
                {
                    var title = node.SelectSingleNode(".//strong");
                    string headline = WebUtility.HtmlDecode((title ?? node).InnerText).Trim();
                    string link = WebUtility.HtmlDecode(node.GetAttributeValue("href", ""));
                    Uri address;
                    if (headline.Length == 0 || !Uri.TryCreate(link, UriKind.Absolute, out address) ||
                        (address.Scheme != "https" && address.Scheme != "http") || !seen.Add(link)) continue;
                    next.Add(new Article { Headline = headline, Link = link });
                    if (next.Count == 6) break;
                }
            }
            if (next.Count == 0) throw new InvalidDataException("뉴스 항목 없음");
            return next;
        }
    }
}
