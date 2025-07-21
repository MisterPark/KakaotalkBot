using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace KakaotalkBot
{
    public static class SmartString
    {
        private static Dictionary<string, Func<string>> format = new Dictionary<string, Func<string>>
        {
            { "닉네임", GetNickname },
        };

        public static string CurrentNickname { get; set; }

        public static string Parse(string input)
        {
            string result = Regex.Replace(input, @"\{(.*?)\}", match =>
            {
                string key = match.Groups[1].Value;
                if (format.TryGetValue(key, out var value))
                {
                    return value.Invoke();
                }
                else
                {
                    return match.Value;
                }
            });

            return result;
        }

        private static string GetNickname()
        {
            return CurrentNickname;
        }
    }
}
