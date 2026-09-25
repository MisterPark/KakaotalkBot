using System;
using System.Collections.Generic;
using System.Linq;
using KakaotalkBot;

// 저장소 단위 시험은 카카오톡 입력·구글 시트를 호출하지 않습니다.
namespace KakaotalkBot
{
    public class User { public long UserId; public int Point; }
    public enum ChatRoomEvent { None, Join, Leave, Kick }
    public class ChatMessage
    {
        public bool IsOwn, IsDirect;
        public ChatRoomEvent RoomEvent;
        public string EventCommand, Message;
        public long AuthorId, ChatId, LogId, SendAt;
    }
}
class DailyQuestTests
{
    static int checks;
    static void Check(bool value, string name)
    { if (!value) throw new Exception(name); checks++; Console.WriteLine("PASS " + name); }
    static void Main()
    {
        long at = new DateTimeOffset(2026, 9, 25, 1, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
        var user = new User { UserId = 9007199254740993L, Point = 100 };
        var store = new DailyQuestStore();
        Check(store.Status(user.UserId, at).Contains("0P / 50P") && store.Rows.Count == 0, "조회만으로 기록·보상 생성 안 함");
        Check(store.Complete(user, "attendance", 10, 1, at).Contains("+10") && user.Point == 110, "출석 퀘스트 지급");
        Check(store.Complete(user, "attendance", 20, 2, at) == "" && user.Point == 110, "다른 방·출석 별칭 공통 중복 방지");
        Check(store.Complete(user, "quiz", 20, 3, at) == "" && user.Point == 110, "첫 정답은 진행도만 증가");
        Check(store.Complete(user, "quiz", 20, 3, at) == "" && store.Status(user.UserId, at).Contains("1/3회"), "같은 정답 로그 중복 집계 방지");
        Check(store.Complete(user, "quiz", 20, 4, at) == "" && user.Point == 110, "두 번째 정답까지 보상 없음");
        var midway = new DailyQuestStore();
        midway.Load(store.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList());
        store = midway;
        Check(store.Status(user.UserId, at).Contains("2/3회"), "정답 진행도 재시작 복원");
        Check(store.Complete(user, "quiz", 10, 3, at).Contains("+20") && user.Point == 130, "세 번째 정답에서 보상 지급·방 간 합산");
        Check(store.Complete(user, "quiz", 10, 4, at) == "", "추가 정답 중복 보상 방지");
        var m = new ChatMessage { AuthorId = user.UserId, ChatId = 10, LogId = 10, SendAt = at, Message = "안녕하세요" };
        m.IsOwn = true; Check(!store.CountChat(user, m, at), "봇 메시지 제외"); m.IsOwn = false;
        m.IsDirect = true; Check(!store.CountChat(user, m, at), "개인 메시지 제외"); m.IsDirect = false;
        m.RoomEvent = ChatRoomEvent.Join; Check(!store.CountChat(user, m, at), "시스템 메시지 제외"); m.RoomEvent = ChatRoomEvent.None;
        m.Message = " /퀘스트"; Check(!store.CountChat(user, m, at), "명령어 제외");
        m.Message = "  "; Check(!store.CountChat(user, m, at), "빈 메시지 제외");
        m.Message = "안녕하세요"; Check(store.CountChat(user, m, at), "첫 채팅 인정");
        Check(!store.CountChat(user, m, at), "같은 로그 재처리 방지");
        m.ChatId = 20; m.LogId = 1; m.Message = "다른 방"; m.SendAt = at + 59;
        Check(!store.CountChat(user, m, at + 59), "방 이동해도 60초 제한 공유");
        m.LogId = 2; m.Message = "안녕하세요"; m.SendAt = at + 60;
        Check(!store.CountChat(user, m, at + 60), "직전 인정 메시지와 동일한 내용 제외");
        m.LogId = 3; m.Message = "새로운 내용";
        Check(store.CountChat(user, m, at + 60), "정확히 60초 뒤 다른 내용 인정");
        var reload = new DailyQuestStore();
        reload.Load(store.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList());
        store = reload;
        Check(store.Rows.Values.Single().UserId == user.UserId && store.Status(user.UserId, at).Contains("2/10"), "64비트 ID와 진행도 재시작 복원");
        Check(store.Complete(user, "quiz", 10, 4, at) == "" && user.Point == 130, "재시작 후 보상 중복 없음");
        Check(!store.CountChat(user, m, at + 60), "재시작 후 채팅 중복 없음");
        for (int i = 2; i < 10; i++)
        { m.LogId = i + 10; m.SendAt = at + i * 60; m.Message = "대화 " + i; if (!store.CountChat(user, m, m.SendAt)) throw new Exception("채팅 집계 누락"); }
        Check(user.Point == 150 && store.Status(user.UserId, at).Contains("50P / 50P"), "채팅 10회와 완주 보상 합계 50P");
        m.LogId++; m.SendAt += 60; m.Message = "추가 대화";
        Check(!store.CountChat(user, m, m.SendAt) && user.Point == 150, "목표 초과 추가 지급 없음");
        Check(store.Rows.Values.Single().Rewards.Count == 4 && store.Rows.Values.Single().Rewards.Values.All(r => r.Key.StartsWith("2026-09-25:")), "퀘스트별 고유 지급 기록");
        long midnight = new DateTimeOffset(2026, 9, 26, 0, 0, 0, TimeSpan.FromHours(9)).ToUnixTimeSeconds();
        Check(DailyQuestStore.Day(midnight - 1) == "2026-09-25" && DailyQuestStore.Day(midnight) == "2026-09-26", "한국 시간 자정 경계");
        Check(store.Status(user.UserId, midnight).Contains("0P / 50P"), "새 날짜 진행도 초기화");
        Check(!store.CountChat(user, m, midnight), "어제 메시지 소급 집계 제외");
        store.Complete(user, "attendance", 10, 200, midnight);
        Check(user.Point == 160 && store.Rows.Count == 2, "다음 날 보상과 과거 기록 보존");
        Console.WriteLine("PASS " + checks + " daily quest checks (no network or native input)");
    }
}
