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
            Console.WriteLine("PASS: quiz memory merge, duplicate suppression, journal recovery, in-flight edits, retry idempotence (no network)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
    }
}
