using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using KakaotalkBot;

internal static class QuizSyncTests
{
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static Quiz Q(string name, string answer = "정답") { return new Quiz { Question = name, Category = "과학", Answer = answer, Difficulty = "하", Hint = "힌트", Explanation = "해설" }; }
    static int Main()
    {
        string directory = Path.Combine(Path.GetTempPath(), "QuizSync-" + Guid.NewGuid().ToString("N"));
        try
        {
            string path = Path.Combine(directory, "changes.json");
            var store = new QuizChangeStore(path);
            var a = Q("로컬 문제"); var b = Q("시트 문제");
            store.Stage(null, a);
            Check(store.Overlay(new List<Quiz> { b }).Count == 2, "메모리 즉시 합치기");
            var restored = new QuizChangeStore(path);
            Check(restored.Count == 1, "재시작 후 미저장 변경 복구");
            var sent = restored.Snapshot();
            var merged = QuizChangeStore.Merge(new List<Quiz> { b }, sent);
            Check(merged.Count == 2, "시트의 새 문제 보존");
            Check(QuizChangeStore.Merge(merged, sent).Count == 2, "응답 유실 후 중복 저장 방지");
            var edited = Q("로컬 문제", "시트 수정 정답");
            var dedup = QuizChangeStore.Merge(new List<Quiz> { edited, b }, sent);
            Check(dedup.Count == 2 && dedup[0].Answer == edited.Answer, "중복은 시트 최신 내용 보존");
            restored.Stage(a, null);
            restored.Confirm(sent);
            Check(restored.Count == 1, "통신 중 추가된 삭제 보존");
            Check(QuizChangeStore.Merge(merged, restored.Snapshot()).Count == 1, "등록 후 삭제 순서");
            var deleteSent = restored.Snapshot();
            restored.Stage(null, Q("로컬 문제", "재등록"));
            restored.Confirm(deleteSent);
            Check(restored.Count == 1 && QuizChangeStore.Merge(new List<Quiz> { b }, restored.Snapshot()).Count == 2, "통신 중 재등록 보존");
            restored.Confirm(restored.Snapshot());
            Check(new QuizChangeStore(path).Count == 0, "반영 완료 기록 제거");
            restored.Stage(a, null);
            Check(QuizChangeStore.Merge(new List<Quiz> { edited }, restored.Snapshot()).Count == 1, "삭제 대기 중 시트 수정 보존");
            Check(QuizChangeStore.Merge(new List<Quiz>(), restored.Snapshot()).Count == 0, "이미 삭제된 문제 재처리");
            var low = Q("최소"); low.AskedCount = 2;
            var tied = Q("동률"); tied.AskedCount = 2;
            var high = Q("다수"); high.AskedCount = 8;
            var pool = new List<Quiz> { high, low, tied };
            Check(Quiz.LeastAsked(pool, null).SequenceEqual(new[] { 1, 2 }), "최소 횟수 동률 후보만 선택");
            Check(Quiz.LeastAsked(pool, new List<int> { 0, 2 }).SequenceEqual(new[] { 2 }), "분류 범위 안에서 최소 선택");
            for (int i = 0; i < 6; i++) { var candidates = Quiz.LeastAsked(pool, null); pool[candidates[0]].AskedCount++; }
            Check(low.AskedCount == 5 && tied.AskedCount == 5 && high.AskedCount == 8, "적게 출제한 문제 우선 순환");
            Check(Quiz.ToCommonSense(a.ToRow().Take(6).Select(Convert.ToString).ToList()).AskedCount == 0, "기존 6열 시트 호환");
            Check(QuizChange.Copy(high).AskedCount == 8, "횟수 복사 및 직렬화");
            var counts = new QuizChangeStore(Path.Combine(directory, "counts.json"));
            var issued = QuizChange.Copy(a); issued.AskedCount = 1;
            counts.Stage(a, issued);
            var sending = counts.Snapshot();
            var updatedRemote = QuizChangeStore.Merge(new List<Quiz> { edited }, sending);
            Check(updatedRemote[0].AskedCount == 1 && updatedRemote[0].Answer == edited.Answer, "횟수 저장 시 시트 본문 수정 보존");
            Check(edited.AskedCount == 0, "원본 스냅샷 불변");
            Check(QuizChangeStore.Merge(updatedRemote, sending)[0].AskedCount == 1, "횟수 재전송 중복 증가 없음");
            Check(QuizChangeStore.Merge(new List<Quiz>(), sending).Count == 0, "시트에서 삭제한 문제를 횟수로 복구하지 않음");
            var twice = QuizChange.Copy(issued); twice.AskedCount = 2;
            counts.Stage(issued, twice); counts.Confirm(sending);
            var recovered = new QuizChangeStore(Path.Combine(directory, "counts.json"));
            Check(recovered.Overlay(updatedRemote)[0].AskedCount == 2, "동기화 중 추가 출제 및 재시작 복구");
            var registered = new QuizChangeStore(Path.Combine(directory, "new.json"));
            registered.Stage(null, a); registered.Stage(a, issued);
            Check(!registered.Snapshot()[0].CountOnly && QuizChangeStore.Merge(new List<Quiz>(), registered.Snapshot())[0].AskedCount == 1, "신규 문제 등록 후 출제도 함께 저장");
            var numbersPath = Path.Combine(directory, "numbers.txt"); var numbers = new QuizNumberStore(numbersPath);
            var numbered = new List<Quiz> { Q("원래 문제"), Q("같은 문장"), Q("같은 문장", "다른 정답") };
            numbers.Assign(numbered);
            Check(numbered.Select(q => q.Number).SequenceEqual(new long[] { 1, 2, 3 }), "기존 중복 문장도 행마다 다른 고유번호 부여");
            numbered.Reverse(); numbers.Assign(numbered);
            Check(numbered.Select(q => q.Number).SequenceEqual(new long[] { 3, 2, 1 }), "목록 재정렬로 고유번호가 바뀌지 않음");
            numbered.RemoveAt(0); numbers = new QuizNumberStore(numbersPath);
            Check(numbers.Next() == 4, "마지막 문제 삭제와 재시작 후에도 번호 재사용 금지");
            numbers.Observe(QuizNumberStore.ReadMaximum(new List<string> { "질문", "분류", "난이도", "답", "힌트", "해설", "출제횟수", "고유번호", "마지막고유번호:99" }));
            Check(numbers.Next() == 100, "시트의 최대 발급 번호를 읽어 다른 PC에서도 삭제 번호 재사용 금지");
            var concurrentNumbers = new QuizNumberStore(numbersPath);
            Check(numbers.Next() == 101 && concurrentNumbers.Next() == 102 && numbers.Next() == 103, "DB 재초기화와 배경 갱신이 겹쳐도 같은 파일의 번호를 중복 발급하지 않음");
            var copied = QuizChange.Copy(numbered[0]);
            Check(copied.Number == numbered[0].Number && Quiz.ToCommonSense(copied.ToRow().Select(Convert.ToString).ToList()).Number == copied.Number, "고유번호 복사 및 시트 직렬화 보존");
            long parsed;
            Check(Quiz.TryParseDeletion("/퀴즈삭제 123", out parsed) && parsed == 123, "고유번호 삭제 명령 파싱");
            foreach (string invalid in new[] { "/퀴즈삭제 0", "/퀴즈삭제 -1", "/퀴즈삭제 1 2", "/퀴즈삭제123", "/퀴즈삭제 과학 | 문제", "/퀴즈삭제 9223372036854775808" })
                Check(!Quiz.TryParseDeletion(invalid, out parsed), "잘못된 고유번호 삭제 거부");
            var first = Q("동일 문제"); first.Number = 200; var secondSame = Q("동일 문제"); secondSame.Number = 201;
            var deletion = new QuizChange { Id = "delete", Before = first };
            Check(QuizChangeStore.Merge(new List<Quiz> { first, secondSame }, new[] { deletion }).Single().Number == 201, "같은 문장이 있어도 지정한 고유번호만 삭제");
            var renumberedText = QuizChange.Copy(first); renumberedText.Question = "시트에서 바꾼 문장";
            Check(QuizChangeStore.Merge(new List<Quiz> { renumberedText }, new[] { deletion }).Count == 1, "삭제 예약 뒤 편집된 문제는 동일 ID라도 보존");
            var renamedCount = QuizChange.Copy(first); renamedCount.AskedCount = 10;
            var renamedMerge = QuizChangeStore.Merge(new List<Quiz> { renumberedText }, new[] { new QuizChange { Id = "count", Before = first, After = renamedCount } });
            Check(renamedMerge.Single().Question == renumberedText.Question && renamedMerge[0].AskedCount == 10, "시트 문장 수정 후에도 고유번호로 출제횟수 반영");
            var legacyChanges = new QuizChangeStore(Path.Combine(directory, "legacy.json")); legacyChanges.Stage(null, Q("기존 미저장 문제"));
            legacyChanges.BindNumbers(numbered, numbers); long assigned = legacyChanges.Snapshot()[0].After.Number;
            Check(assigned > 0 && new QuizChangeStore(Path.Combine(directory, "legacy.json")).Snapshot()[0].After.Number == assigned, "업데이트 전 미저장 등록에도 영구 고유번호 배정");
            var many = Enumerable.Range(1, 100).Select(i => new Quiz { Number = i, Category = "과학", Difficulty = "하", Question = "문제" + i + new string('가', 90), Answer = "숨긴 정답" }).ToList();
            var pages = Quiz.ListMessages(many);
            Check(pages.Count > 1 && pages.All(p => p.Length < 3100) && many.All(q => pages.Any(p => p.Contains(q.Number + ". [과학 · 하] " + q.Question))) && !string.Join("", pages).Contains("숨긴 정답"), "전체 문제를 고유번호와 함께 누락 없이 메시지 분할");
            bool duplicate = false; try { numbers.Assign(new List<Quiz> { first, QuizChange.Copy(first) }); } catch (InvalidOperationException) { duplicate = true; }
            Check(duplicate, "중복된 시트 고유번호를 자동으로 바꾸거나 잘못 삭제하지 않음");
            Console.WriteLine("PASS: quiz memory merge, duplicate suppression, journal recovery, in-flight edits, retry idempotence (no network)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
    }
}
