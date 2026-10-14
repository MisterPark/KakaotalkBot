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
            Console.WriteLine("PASS: quiz memory merge, duplicate suppression, journal recovery, in-flight edits, retry idempotence (no network)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
    }
}
