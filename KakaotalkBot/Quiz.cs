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
            if (command != null && (command.StartsWith("/퀴즈등록", StringComparison.Ordinal) || command.StartsWith("/퀴즈삭제", StringComparison.Ordinal))) return false;
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

        public const string DeletionHelp = "/퀴즈삭제 분류 | 문제 전체 문장\n예: /퀴즈삭제 과학 | 물의 화학식은?";

        public static bool TryParseDeletion(string text, out string category, out string question)
        {
            category = question = null;
            const string name = "/퀴즈삭제";
            if (text == null || !text.StartsWith(name, StringComparison.Ordinal) ||
                text.Length <= name.Length || !char.IsWhiteSpace(text[name.Length])) return false;
            var parts = text.Substring(name.Length).Split('|');
            if (parts.Length != 2) return false;
            category = parts[0].Trim(); question = parts[1].Trim();
            if (category.Length == 0 || category.Length > 30 || question.Length == 0 || question.Length > 1000) return false;
            foreach (char c in category + question) if (char.IsControl(c)) return false;
            return true;
        }

        public static int FindDeletionRow(List<List<string>> rows, string category, string question)
        {
            int found = -1;
            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                if (row.Count < 2 || !string.Equals(row[0].Trim(), question, StringComparison.Ordinal) ||
                    !string.Equals(row[1].Trim(), category, StringComparison.Ordinal)) continue;
                if (found >= 0) throw new InvalidOperationException("같은 분류와 문제의 행이 여러 개입니다. 시트에서 대상을 확인해 주세요.");
                found = i;
            }
            if (found < 0) throw new InvalidOperationException("일치하는 문제가 없습니다. 분류와 문제 전체 문장을 확인해 주세요.");
            return found;
        }

        public const string RegistrationHelp = "/퀴즈등록 분류 | 난이도 | 문제 | 정답 | 힌트 | 해설\n예: /퀴즈등록 과학 | 하 | 물의 화학식은? | H2O | 알파벳과 숫자 | 수소 두 개와 산소 한 개";

        public static bool TryParseRegistration(string text, out Quiz quiz, out string error)
        {
            quiz = null;
            error = RegistrationHelp;
            const string name = "/퀴즈등록";
            if (text == null || !text.StartsWith(name, StringComparison.Ordinal) ||
                text.Length <= name.Length || !char.IsWhiteSpace(text[name.Length])) return false;
            string[] parts = text.Substring(name.Length).Split('|');
            if (parts.Length != 6) return false;
            int[] limits = { 30, 20, 1000, 200, 500, 1500 };
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].Trim();
                if (parts[i].Length == 0 || parts[i].Length > limits[i])
                { error = "모든 항목을 입력해 주세요. 분류 30자, 난이도 20자, 문제 1000자, 정답 200자, 힌트 500자, 해설 1500자까지 가능합니다."; return false; }
                foreach (char c in parts[i])
                    if (char.IsControl(c)) { error = "각 항목은 줄바꿈 없이 입력해 주세요."; return false; }
            }
            if (parts[0].StartsWith("/", StringComparison.Ordinal))
            { error = "분류에는 명령어의 /를 넣지 마세요."; return false; }
            quiz = new Quiz { Category = parts[0], Difficulty = parts[1], Question = parts[2],
                Answer = parts[3], Hint = parts[4], Explanation = parts[5] };
            error = null;
            return true;
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
