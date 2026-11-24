using System;
using System.Collections.Generic;
using System.Linq;

namespace KakaotalkBot
{
    internal static class BannedListLoading
    {
        internal const int TimeoutMilliseconds = 10000;
        private const int StableMilliseconds = 500;

        // 창 생성과 서버 응답은 별개입니다. 실제 창은 0명으로 생성된 후 항목이 들어옵니다.
        // 빈 트리 자체에는 로딩 완료 표시가 없으므로, 0명을 조회 성공으로 반환하지 않습니다.
        internal static List<BannedMember> Wait(Func<List<BannedMember>> read, Action guard, Func<long> elapsed, Action<int> sleep)
        {
            List<BannedMember> previous = null;
            long stableSince = 0;
            while (elapsed() < TimeoutMilliseconds)
            {
                guard();
                var rows = read();
                guard();
                long now = elapsed();
                if (now >= TimeoutMilliseconds) break;
                if (rows == null || rows.Count == 0)
                    previous = null;
                else
                {
                    // 읽는 도중 창·사용자·이름·일시가 바뀌면 다시 안정화 시간을 기다립니다.
                    var ordered = rows.OrderBy(r => r.UserId).ThenBy(r => r.KickedAt).ToList();
                    if (previous == null || !Same(previous, ordered))
                    {
                        previous = ordered.Select(r => new BannedMember { ListWindow = r.ListWindow, UserId = r.UserId,
                            KickedAt = r.KickedAt, Name = r.Name }).ToList();
                        stableSince = now;
                    }
                    else if (now - stableSince >= StableMilliseconds)
                        return rows;
                }
                sleep(50);
            }
            throw new TimeoutException("내보낸 목록의 로딩 완료를 확인하지 못했습니다. 잠시 후 /내보낸목록으로 다시 확인하세요.");
        }

        private static bool Same(List<BannedMember> first, List<BannedMember> second)
        {
            if (first.Count != second.Count) return false;
            for (int i = 0; i < first.Count; i++)
                if (first[i].ListWindow != second[i].ListWindow || first[i].UserId != second[i].UserId ||
                    first[i].KickedAt != second[i].KickedAt || first[i].Name != second[i].Name) return false;
            return true;
        }
    }
}
