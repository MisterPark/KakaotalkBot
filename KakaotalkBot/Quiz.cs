using System;
using System.Collections.Generic;

namespace KakaotalkBot
{
    public class Quiz
    {
        public string Question = string.Empty;
        public string Category = string.Empty;
        public string Difficulty = string.Empty;
        public string Answer = string.Empty;
        public string Hint = string.Empty;
        public string Explanation = string.Empty;

        public static bool IsQuizCommand(string command)
        {
            return command == "/퀴즈" || command == "/퀴즈목록" ||
                (command != null && command.StartsWith("/", StringComparison.Ordinal) &&
                 command.EndsWith("퀴즈", StringComparison.Ordinal) && command.IndexOfAny(new[] { '\r', '\n' }) < 0);
        }

        // 갱신할 때 한 번만 분류별 원본 인덱스를 만들고 출제 시 재사용합니다.
        internal static Dictionary<string, List<int>> BuildCategoryIndex(List<Quiz> quizzes)
        {
            var result = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < quizzes.Count; i++)
            {
                string category = (quizzes[i].Category ?? "").Trim();
                if (category.Length == 0) continue;
                List<int> indexes;
                if (!result.TryGetValue(category, out indexes)) result.Add(category, indexes = new List<int>());
                indexes.Add(i);
            }
            return result;
        }

        public List<object> ToRow()
        {
            return new List<object>() { Question, Category, Difficulty, Answer, Hint, Explanation };
        }

        public static Quiz ToCommonSense(List<string> list)
        {
            Quiz cs = new Quiz();
            cs.Question = list[0];
            cs.Category = list[1];
            cs.Difficulty = list[2];
            cs.Answer = list[3];
            cs.Hint = list[4];
            cs.Explanation = list[5];

            return cs;
        }
    }
}
