using System.Collections.Generic;

namespace KakaotalkBot
{
    public class CommonSense
    {
        public string Question = string.Empty;
        public string Category = string.Empty;
        public string Difficulty = string.Empty;
        public string Answer = string.Empty;
        public string Hint = string.Empty;
        public string Explanation = string.Empty;

        public List<object> ToRow()
        {
            return new List<object>() { Question, Category, Difficulty, Answer, Hint, Explanation };
        }

        public static CommonSense ToCommonSense(List<string> list)
        {
            CommonSense cs = new CommonSense();
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
