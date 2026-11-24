using System;
using System.Collections.Generic;
using System.Linq;
using KakaotalkBot;

internal static class BannedListLoadingTests
{
    private static int passed;
    private static void Check(bool ok, string message)
    { if (!ok) throw new Exception(message); passed++; Console.WriteLine("PASS " + message); }
    private static List<BannedMember> Rows(int count, int window = 1, long first = 100)
    {
        return Enumerable.Range(0, count).Select(i => new BannedMember { UserId = first + i, KickedAt = 1700000000 + i,
            Name = "사용자 " + i, ListWindow = new IntPtr(window) }).ToList();
    }
    private sealed class Run
    {
        internal long Now;
        internal int Guards;
        internal List<BannedMember> Wait(Func<long, List<BannedMember>> read, Action guard = null)
        {
            return BannedListLoading.Wait(() => read(Now), () => { Guards++; if (guard != null) guard(); }, () => Now, milliseconds => Now += milliseconds);
        }
    }
    private static bool Fails(Action action, string text)
    {
        try { action(); return false; }
        catch (InvalidOperationException error) { return error.Message.Contains(text); }
        catch (TimeoutException error) { return error.Message.Contains(text); }
    }
    private static int Main()
    {
        try
        {
            var run = new Run();
            var result = run.Wait(t => t < 71 ? Rows(0) : Rows(6));
            Check(result.Count == 6 && run.Now >= 600, "live regression: initial zero followed by six members is not published as empty");

            run = new Run(); result = run.Wait(t => t < 7000 ? Rows(0) : Rows(6));
            Check(result.Count == 6 && run.Now >= 7500, "slow response waits for actual members instead of assuming an empty list");

            run = new Run(); result = run.Wait(t => t < 100 ? null : t < 300 ? Rows(0) : Rows(6));
            Check(result.Count == 6 && run.Now >= 800, "window creation and empty initialization can precede populated rows");

            run = new Run(); result = run.Wait(t => t < 300 ? Rows(2) : Rows(6));
            Check(result.Count == 6 && run.Now >= 800, "partially populated list must stabilize again after additional members arrive");

            run = new Run(); result = run.Wait(t => t < 300 ? Rows(6) : Rows(6, 2));
            Check(result[0].ListWindow == new IntPtr(2) && run.Now >= 800, "a replaced window cannot reuse stability of the previous window");

            run = new Run(); result = run.Wait(t => t < 300 ? Rows(6) : Rows(6, 1, 200));
            Check(result[0].UserId == 200 && run.Now >= 800, "same count with different members must stabilize again");

            run = new Run(); result = run.Wait(t => { var rows = Rows(6); if (t >= 300) rows[0].Name = "변경된 이름"; return rows; });
            Check(result[0].Name == "변경된 이름" && run.Now >= 800, "late name update is included in the returned snapshot");

            run = new Run(); result = run.Wait(t => { var rows = Rows(6); if (t >= 300) rows[0].KickedAt++; return rows; });
            Check(result[0].KickedAt == 1700000001 && run.Now >= 800, "changed removal timestamp restarts stabilization");

            run = new Run(); result = run.Wait(t => t == 300 ? null : Rows(6));
            Check(result.Count == 6 && run.Now >= 850, "transient read failure interrupts stabilization");

            run = new Run(); result = run.Wait(t => t == 300 ? Rows(0) : Rows(6));
            Check(result.Count == 6 && run.Now >= 850, "empty refresh interrupts stabilization");

            run = new Run(); result = run.Wait(t => { var rows = Rows(6); if ((t / 50) % 2 == 0) rows.Reverse(); return rows; });
            Check(result.Count == 6 && run.Now == 500, "ordering alone does not delay a stable set of members");

            run = new Run(); var shared = Rows(6);
            result = run.Wait(t => { if (t >= 300) shared[0].Name = "변경"; return shared; });
            Check(result[0].Name == "변경" && run.Now >= 800, "snapshot comparison detects in-place mutation by a reader");

            run = new Run(); bool recorded = false;
            Check(Fails(() => { run.Wait(t => Rows(0)); recorded = true; }, "로딩 완료") && !recorded && run.Now == 10000,
                "unconfirmed empty list times out without a successful result that would clear saved bans");

            run = new Run();
            Check(Fails(() => run.Wait(t => null), "로딩 완료") && run.Now == 10000, "missing window has a bounded wait");

            run = new Run();
            Check(Fails(() => run.Wait(t => t < 9800 ? Rows(0) : Rows(6)), "로딩 완료"), "deadline cannot return a snapshot still changing");

            run = new Run();
            Check(Fails(() => run.Wait(t => Rows(6, 1, t / 50 % 2 == 0 ? 100 : 200)), "로딩 완료"), "continuously changing list never produces an arbitrary result");

            run = new Run();
            Check(Fails(() => run.Wait(t => Rows(6), () => { if (run.Now >= 200) throw new InvalidOperationException("cancelled"); }), "cancelled") && run.Now == 200,
                "room or permission change cancels the load promptly");

            run = new Run(); bool revoked = false;
            Check(Fails(() => run.Wait(t => { if (t >= 500) revoked = true; return Rows(6); }, () => { if (revoked) throw new InvalidOperationException("revoked"); }), "revoked"),
                "permission is checked again before accepting a completed read");

            Console.WriteLine("Banned-list loading checks passed: " + passed); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
