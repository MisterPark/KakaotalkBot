using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

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
            string url = "https://news.naver.com/breakingnews/section/100/264";
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");

            // 동기적으로 HTML 가져오기
            var html = httpClient.GetStringAsync(url).GetAwaiter().GetResult();

            int idx = html.IndexOf("SECTION_ARTICLE_LIST_FOR_LATEST");
            string sub = html.Substring(idx);
            string[] split = sub.Split(new string[] { "<", ">" }, StringSplitOptions.RemoveEmptyEntries);
            List<string> links = split.Where(x => x.StartsWith("a href=")).Select(x =>
            {
                string link = x.Split(new char[] { '\"' })[1];
                return link;
            }).ToList();

            List<string> top6 = new List<string>();
            for (int i = 0; i < 16; i++)
            {
                if (i % 3 == 0)
                {
                    top6.Add(links[i]);
                }
            }

            List<string> headlines = sub.Split(new string[] { "<" }, StringSplitOptions.RemoveEmptyEntries)
                .Where(x => x.StartsWith("strong class=\"sa_text_strong\">"))
                .Select(x =>
                {
                    string replaced = x.Replace("strong class=\"sa_text_strong\">", "");
                    return WebUtility.HtmlDecode(replaced);
                }).ToList();

            articles.Clear();
            for (int i = 0; i < 6; i++)
            {
                Article article = new Article();
                article.Headline =  headlines[i];
                article.Link = links[i];
                articles.Add(article);
            }
        }
    }
}
