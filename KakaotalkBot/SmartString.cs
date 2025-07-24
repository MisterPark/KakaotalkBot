using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace KakaotalkBot
{
    public enum ParseResult
    {
        Success,
        InvalidFormat,
    }
    public static class SmartString
    {
        private static Dictionary<string, Func<string>> format = new Dictionary<string, Func<string>>
        {
            { "닉네임", GetNickname },
        };

        public static string CurrentNickname { get; set; }

        public static ParseResult ParseCommand(string input, out string[] format)
        {
            format = new string[3];
            string[] parts = input.Split(new char[]{ ' '},StringSplitOptions.None);

            if (parts.Length < 3)
            {
                return ParseResult.InvalidFormat;
            }

            // 맨 뒤가 숫자인지 확인
            int form = 1;
            if (!int.TryParse(parts[parts.Length -1], out int number))
            {
                form = 0;
            }

            string command = parts[0]; // /kick

            // 닉네임은 [1]부터 [^2]까지 이어붙인 것
            string nickname = "";
            for (int i = 1; i < parts.Length - form; i++)
            {
                if (i > 1) nickname += " ";
                nickname += parts[i];
            }

            format[0] = command;
            format[1] = nickname;
            format[2] = number.ToString();

            return ParseResult.Success;

        }
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
