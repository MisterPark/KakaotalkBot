using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

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
        public long AskedCount;
        public long Number;

        // 전체 정렬 없이 최소 출제 횟수인 원본 인덱스만 모읍니다.
        internal static List<int> LeastAsked(List<Quiz> quizzes, List<int> indexes)
        {
            var candidates = new List<int>();
            long minimum = long.MaxValue;
            int count = indexes == null ? quizzes.Count : indexes.Count;
            for (int i = 0; i < count; i++)
            {
                int index = indexes == null ? i : indexes[i];
                long value = quizzes[index].AskedCount;
                if (value < minimum) { minimum = value; candidates.Clear(); }
                if (value == minimum) candidates.Add(index);
            }
            return candidates;
        }

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

        public const string DeletionHelp = "/퀴즈삭제 고유번호\n예: /퀴즈삭제 123\n번호는 /퀴즈목록에서 확인하세요.";

        public static bool TryParseDeletion(string text, out long number)
        {
            number = 0; const string command = "/퀴즈삭제";
            return text != null && text.StartsWith(command, StringComparison.Ordinal) && text.Length > command.Length && char.IsWhiteSpace(text[command.Length]) &&
                long.TryParse(text.Substring(command.Length).Trim(), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out number) && number > 0;
        }

        internal static List<string> ListMessages(IEnumerable<Quiz> quizzes)
        {
            var rows = quizzes.OrderBy(q => q.Number).ToArray();
            if (rows.Length == 0) return new List<string> { "[퀴즈목록]\n등록된 퀴즈가 없습니다." };
            if (rows.Any(q => q.Number <= 0) || rows.Select(q => q.Number).Distinct().Count() != rows.Length)
                throw new InvalidOperationException("퀴즈 고유번호를 준비 중입니다. 콘텐츠 갱신 상태를 확인해 주세요.");
            var chunks = new List<string>(); var body = new System.Text.StringBuilder();
            foreach (var quiz in rows)
            {
                string line = quiz.Number + ". [" + quiz.Category + " · " + quiz.Difficulty + "] " + quiz.Question + "\n";
                if (body.Length > 0 && line.Length <= 3000 && body.Length + line.Length > 3000) { chunks.Add(body.ToString().TrimEnd()); body.Clear(); }
                // 시트에서 직접 넣은 긴 문장도 생략하지 않고 나눠 전송합니다.
                while (line.Length > 0)
                {
                    int count = Math.Min(3000 - body.Length, line.Length);
                    if (count < line.Length && count > 0 && char.IsHighSurrogate(line[count - 1])) count--;
                    if (count == 0) { chunks.Add(body.ToString().TrimEnd()); body.Clear(); continue; }
                    body.Append(line.Substring(0, count)); line = line.Substring(count);
                    if (body.Length >= 2999 || line.Length > 0) { chunks.Add(body.ToString().TrimEnd()); body.Clear(); }
                }
            }
            if (body.Length > 0) chunks.Add(body.ToString().TrimEnd());
            return chunks.Select((chunk, i) => "[퀴즈목록 " + (i + 1) + "/" + chunks.Count + " · 전체 " + rows.Length + "개]\n" + chunk).ToList();
        }

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
            return new List<object>() { Question, Category, Difficulty, Answer, Hint, Explanation, AskedCount, Number > 0 ? Number.ToString(System.Globalization.CultureInfo.InvariantCulture) : "" };
        }

        public static Quiz ToCommonSense(List<string> list)
        {
            Quiz cs = new Quiz();
            cs.Question = list[0];
            cs.Category = list[1];
            cs.Difficulty = list[2];
            cs.Answer = list[3];
            cs.Hint = list[4];
            cs.Explanation = list.Count > 5 ? list[5] : "";
            if (list.Count > 6 && !string.IsNullOrWhiteSpace(list[6]) &&
                (!long.TryParse(list[6], out cs.AskedCount) || cs.AskedCount < 0))
                throw new InvalidDataException("퀴즈 출제횟수는 0 이상의 정수여야 합니다.");
            if (list.Count > 7 && !string.IsNullOrWhiteSpace(list[7]) &&
                (!long.TryParse(list[7], out cs.Number) || cs.Number <= 0)) throw new InvalidDataException("퀴즈 고유번호는 양의 정수여야 합니다.");

            return cs;
        }
    }
    // 한 시트에 한 봇이 쓰는 기존 운영 방식. 삭제/재시작 후에도 번호를 다시 쓰지 않습니다.
    internal sealed class QuizNumberStore
    {
        private readonly string path;
        private static readonly Dictionary<string, object> gates = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly object gate;
        private long last;
        internal QuizNumberStore(string path)
        {
            this.path = Path.GetFullPath(path);
            lock (gates) { if (!gates.TryGetValue(this.path, out gate)) gates.Add(this.path, gate = new object()); }
            lock (gate) Refresh();
        }
        private void Refresh()
        {
            if (!File.Exists(path)) return;
            long stored;
            if (!long.TryParse(File.ReadAllText(path), out stored) || stored < 0) throw new InvalidDataException("퀴즈 번호 기록이 잘못되었습니다.");
            last = Math.Max(last, stored);
        }
        internal long Last { get { lock (gate) { Refresh(); return last; } } }
        internal static string MaximumCell(long number) { return "마지막고유번호:" + number.ToString(System.Globalization.CultureInfo.InvariantCulture); }
        internal static long ReadMaximum(List<string> header)
        {
            if (header.Count < 9) return 0;
            const string prefix = "마지막고유번호:"; long value;
            if (!header[8].StartsWith(prefix, StringComparison.Ordinal) || !long.TryParse(header[8].Substring(prefix.Length), out value) || value < 0)
                throw new InvalidDataException("퀴즈 번호 기록 열이 변경되었습니다.");
            return value;
        }
        private void Save(long value)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)); string temp = path + ".tmp";
            File.WriteAllText(temp, value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path);
            last = value;
        }
        internal void Observe(long number) { lock (gate) { Refresh(); if (number > last) Save(number); } }
        internal long Next() { lock (gate) { Refresh(); if (last == long.MaxValue) throw new InvalidOperationException("퀴즈 번호 범위를 초과했습니다."); Save(last + 1); return last; } }
        internal void Assign(List<Quiz> quizzes)
        {
            lock (gate)
            {
                Refresh();
                var ids = quizzes.Where(q => q.Number > 0).Select(q => q.Number).ToArray();
                if (ids.Distinct().Count() != ids.Length) throw new InvalidOperationException("퀴즈 고유번호가 중복되어 있습니다. 시트의 고유번호 열을 확인하세요.");
                long maximum = Math.Max(last, ids.Length == 0 ? 0 : ids.Max());
                var missing = quizzes.Where(q => q.Number == 0).ToArray();
                if (maximum > long.MaxValue - missing.Length) throw new InvalidOperationException("퀴즈 번호 범위를 초과했습니다.");
                if (maximum + missing.Length > last) Save(maximum + missing.Length);
                foreach (var quiz in missing) quiz.Number = ++maximum;
            }
        }
    }
    // 이 목록은 봇 처리 스레드만 수정하며, 통신 작업에는 복사본만 전달합니다.
    internal sealed class QuizChange
    {
        public string Id;
        public Quiz Before, After;
        internal bool CountOnly { get { return Before != null && After != null && SameValue(Before, After); } }
        internal Quiz Target { get { return After ?? Before; } }
        internal static bool SameKey(Quiz a, Quiz b)
        { return a != null && b != null && (a.Number > 0 && b.Number > 0 ? a.Number == b.Number : SameText(a, b)); }
        internal static bool SameText(Quiz a, Quiz b)
        { return a != null && b != null && a.Category.Trim() == b.Category.Trim() && a.Question.Trim() == b.Question.Trim(); }
        internal static bool SameValue(Quiz a, Quiz b)
        { return a == null || b == null ? a == b : a.ToRow().Take(6).SequenceEqual(b.ToRow().Take(6)); }
        internal static Quiz Copy(Quiz q)
        { return q == null ? null : Quiz.ToCommonSense(q.ToRow().Select(Convert.ToString).ToList()); }
        internal QuizChange Copy()
        { return new QuizChange { Id = Id, Before = Copy(Before), After = Copy(After) }; }
    }

    internal sealed class QuizChangeStore
    {
        private readonly string path;
        private List<QuizChange> changes = new List<QuizChange>();
        internal int Count { get { return changes.Count; } }
        internal QuizChangeStore(string path)
        {
            this.path = path;
            if (File.Exists(path))
            {
                changes = Newtonsoft.Json.JsonConvert.DeserializeObject<List<QuizChange>>(File.ReadAllText(path));
                if (changes == null || changes.Any(c => c == null || c.Target == null || string.IsNullOrEmpty(c.Id)))
                    throw new InvalidDataException("미저장 퀴즈 변경 기록이 잘못되었습니다.");
            }
        }
        internal QuizChange[] Snapshot() { return changes.Select(c => c.Copy()).ToArray(); }
        internal void BindNumbers(List<Quiz> source, QuizNumberStore numbers)
        {
            var next = changes.Select(c => c.Copy()).ToList(); bool changed = false;
            foreach (var c in next)
            {
                if (c.Target.Number > 0) { numbers.Observe(c.Target.Number); continue; }
                var matches = source.Where(q => QuizChange.SameText(q, c.Target)).ToArray();
                if (matches.Length > 1) throw new InvalidOperationException("이전 퀴즈 변경 기록의 대상이 중복되어 고유번호를 연결하지 못했습니다.");
                long number = matches.Length == 1 ? matches[0].Number : numbers.Next();
                if (c.Before != null) c.Before.Number = number;
                if (c.After != null) c.After.Number = number;
                changed = true;
            }
            if (changed) Save(next);
        }
        private void Save(List<QuizChange> next)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            File.WriteAllText(temp, Newtonsoft.Json.JsonConvert.SerializeObject(next, Newtonsoft.Json.Formatting.Indented).Replace("\r\n", "\n").Replace("\n", "\r\n"));
            if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path);
            changes = next;
        }
        internal void Stage(Quiz before, Quiz after)
        {
            var next = changes.Select(c => c.Copy()).ToList();
            var old = next.FirstOrDefault(c => QuizChange.SameKey(c.Target, after ?? before));
            if (old != null) next.Remove(old);
            next.Add(new QuizChange { Id = Guid.NewGuid().ToString("N"), Before = QuizChange.Copy(old == null ? before : old.Before ?? (after == null ? before : null)), After = QuizChange.Copy(after) });
            Save(next);
        }
        internal void Confirm(QuizChange[] sent)
        {
            var ids = new HashSet<string>(sent.Select(c => c.Id));
            // 통신 중 같은 문제에 추가된 변경은 다른 ID이므로 그대로 남깁니다.
            var next = changes.Where(c => !ids.Contains(c.Id)).Select(c => c.Copy()).ToList();
            foreach (var c in next)
            {
                var prior = sent.FirstOrDefault(p => QuizChange.SameKey(p.Target, c.Target));
                if (prior != null) c.Before = QuizChange.Copy(prior.After);
            }
            Save(next);
        }
        internal List<Quiz> Overlay(List<Quiz> source)
        {
            var result = new List<Quiz>(source);
            foreach (var c in changes)
            {
                if (c.CountOnly)
                {
                    result = Merge(result, new[] { c });
                    continue;
                }
                result.RemoveAll(q => QuizChange.SameKey(q, c.Target));
                if (c.After != null) result.Add(QuizChange.Copy(c.After));
            }
            return result;
        }
        // 서버 상태가 이미 목표와 같으면 이전 요청은 반영된 것으로 간주합니다.
        internal static List<Quiz> Merge(List<Quiz> remote, QuizChange[] pending)
        {
            var result = remote.Select(QuizChange.Copy).ToList();
            foreach (var c in pending)
            {
                var matches = result.Where(q => QuizChange.SameKey(q, c.Target)).ToList();
                if (c.Before == null && c.After != null && c.After.Number > 0 && matches.Any(q => !QuizChange.SameText(q, c.After)))
                    throw new InvalidOperationException("동일한 퀴즈 고유번호가 다른 문제에 사용되어 동기화를 중단했습니다.");
                if (c.After != null)
                {
                    // 시트에 같은 분류·문제가 있으면 중복 등록하지 않고 시트의 최신 내용을 사용합니다.
                    if (matches.Count == 0 && !c.CountOnly) result.Add(QuizChange.Copy(c.After));
                    // 출제 기록의 재전송은 누적 덧셈 대신 목표값 비교로 중복 집계를 막습니다.
                    foreach (var match in matches) match.AskedCount = Math.Max(match.AskedCount, c.After.AskedCount);
                }
                else
                {
                    // 삭제 예약 후 시트에서 수정된 문제는 원문이 달라졌으므로 보존합니다.
                    result.RemoveAll(q => QuizChange.SameKey(q, c.Target) && QuizChange.SameValue(q, c.Before));
                }
            }
            return result;
        }
    }

}
