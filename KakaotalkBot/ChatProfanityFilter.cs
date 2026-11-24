using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace KakaotalkBot
{
    // 저장 원본·검색어·사용자 ID는 바꾸지 않고 전송할 표시 문자열만 가립니다.
    internal static class ChatProfanityFilter
    {
        private const string Gap = @"[\p{Zs}\t\p{Cf}\p{P}]{0,8}";
        private static string Spaced(string word)
        {
            return string.Join(Gap, word.Select(c => c >= 'a' && c <= 'z'
                ? "[" + c + (char)(c + 0xFEE0) + "]" : Regex.Escape(c.ToString())));
        }
        private static readonly Regex Pattern = CreatePattern();
        private static Regex CreatePattern()
        {
            var korean = new[] {
                "씨발", "씨팔", "씨벌", "시팔", "시벌", "씹할", "씨바", "시바새끼",
                "개새끼", "개새기", "개새키", "개색끼", "개색기", "개세끼", "개세키", "개쌔끼", "개쌔키",
                "씹새끼", "씹새", "씹년", "씹놈", "미친놈", "미친년", "미친새끼",
                "병신", "븅신", "빙신", "등신", "지랄", "지럴", "염병", "옘병",
                "좆", "좇같", "좃같", "존나", "존내", "좇나", "좃나",
                "니미", "니애미", "니애비", "느금마", "느개비", "애미뒤진", "애비뒤진",
                "개지랄", "개소리", "개년", "썅년", "썅놈", "쌍년", "쌍놈", "닥쳐", "꺼져", "엿먹어"
            };
            var english = new[] { "motherfucker", "motherfuckers", "motherfucking", "fuck", "fucks", "fucker", "fuckers", "fucking", "fucked",
                "shit", "shits", "shitty", "bullshit", "bitch", "bitches", "bitching", "asshole", "assholes", "bastard", "bastards", "cunt", "cunts", "dickhead" };
            string latin = "a-zａ-ｚ";
            string words = string.Join("|", korean.OrderByDescending(w => w.Length).Select(Spaced));
            // '시발점/시발역/시발지/시발차'와 '새끼손가락/새끼줄', 영단어의 일부는 가리지 않습니다.
            words += "|" + Spaced("시발") + "(?!점|역|지|차)";
            words += "|" + Spaced("새끼") + "(?!손가락|줄)";
            words += "|(?<![" + latin + "])(?:" + string.Join("|", english.OrderByDescending(w => w.Length).Select(Spaced)) + ")(?![" + latin + "])";
            words += @"|(?<![\p{L}\p{N}])(?:" + string.Join("|", new[] { "ㅅㅂ", "ㅆㅂ", "ㅂㅅ", "ㅈㄴ", "ㅈㄹ", "ㄲㅈ" }.Select(Spaced)) + @")(?![\p{L}\p{N}])";
            return new Regex(words, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled, TimeSpan.FromMilliseconds(200));
        }

        internal static string Mask(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            try { return Pattern.Replace(text, match => new string('*', match.Length)); }
            // 필터 검사에 실패한 원문을 그대로 채팅에 내보내지 않습니다.
            catch (RegexMatchTimeoutException) { return "[안내]\n응답 내용을 확인하지 못했습니다. 잠시 후 다시 시도해 주세요."; }
        }

        internal static bool NeedsMasking(string text) { return Mask(text) != text; }
    }
}
