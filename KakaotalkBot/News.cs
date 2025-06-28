using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using HtmlAgilityPack;

namespace KakaotalkBot
{
    public static class News
    {
        private static List<Article> articles = new List<Article>();
        public static string PoliticsTop6
        {
            get
            {
                int num = 1;
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("[정치뉴스 TOP6]");
                sb.AppendLine();
                foreach (var article in articles)
                {
                    sb.AppendLine($"{num}.{article.Headline}");
                    sb.AppendLine($"{article.Link}");
                    num++;
                }

                return sb.ToString();
            }
        }

        public static void Update()
        {
            UpdatePoliticsTop6();
        }

        private static void UpdatePoliticsTop6()
        {
            string url = "https://news.naver.com/breakingnews/section/100/269";
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");

            // 동기적으로 HTML 가져오기
            var html = httpClient.GetStringAsync(url).GetAwaiter().GetResult();

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            articles.Clear();
            for (int i = 0; i < 6; i++)
            {
                var node = doc.DocumentNode.SelectSingleNode($"/html/body/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div[1]/ul/li[{i+1}]/div/div/div[2]/a");
                var strong = doc.DocumentNode.SelectSingleNode($"/html/body/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div[1]/ul/li[{i+1}]/div/div/div[2]/a/strong");

                string link = string.Empty;
                string headline = string.Empty;

                if (node != null)
                {
                    link = node.GetAttributeValue("href", null);
                }
                if (strong != null)
                {
                    headline = WebUtility.HtmlDecode(strong.InnerText);
                }

                Article article = new Article();
                article.Headline = headline;
                article.Link = link;
                articles.Add(article);
            }
        }
    }
}
