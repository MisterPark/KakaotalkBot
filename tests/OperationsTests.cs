using System;
using System.Linq;
using System.Collections.Generic;
using KakaotalkBot;

class OperationsTests
{
    static int checks;
    static void Check(bool result, string name) { if (!result) throw new Exception(name); checks++; }
    static long At(string value) { return DateTimeOffset.Parse(value).ToUnixTimeSeconds(); }
    static OperationsStore Reload(OperationsStore value)
    { var next = new OperationsStore(); next.Load(value.ToRows().Select(r => r.Select(Convert.ToString).ToList()).ToList()); return next; }
    static int Main()
    {
        try
        {
            long before = At("2026-09-30T14:59:59Z"), boundary = before + 1;
            var user = new User(9007199254740993L) { Point = 123, Popularity = 12, Experience = 99 };
            var other = new User(22) { Point = 456 };
            var users = new[] { user, other };
            var store = new OperationsStore();
            Check(OperationsStore.Month(before) == "2026-09" && OperationsStore.Month(boundary) == "2026-10", "한국 자정 경계");
            Check(store.AdvanceMonth(users, before) && user.Point == 123, "최초 도입 기존 포인트 유지");
            store.ResetPending = false;
            Check(!store.AdvanceMonth(users, before), "같은 월 중복 처리 없음");
            Check(store.AdvanceMonth(users, boundary) && users.All(u => u.Point == 0), "모든 사용자 월 초기화");
            Check(user.Popularity == 12 && user.Experience == 99, "포인트 외 정보 유지");
            Check(store.Rows["reset:2026-10:" + user.UserId].Count == 123, "이전 잔액 보관");
            user.Point = 9;
            Check(store.AdvanceMonth(users, boundary + 1) && user.Point == 9, "저장 실패 재시도는 재차 차감하지 않음");
            store = Reload(store);
            Check(!store.AdvanceMonth(users, boundary + 2) && user.Point == 9, "재시작 월 초기화 중복 없음");
            Check(!store.AdvanceMonth(users, before), "시계 역행 초기화 금지");
            Check(store.AdvanceMonth(users, At("2027-01-03T00:00:00Z")) && user.Point == 0, "미실행 기간 다음 실행에 초기화");
            var message = new ChatMessage { ChatId = 10, AuthorId = user.UserId, LogId = 100, SendAt = boundary, Message = "채팅", Nickname = "이름" };
            store.CountMessage(message); store.CountMessage(message);
            Check(store.Rows["monthly:2026-10:10:" + user.UserId].Count == 1, "채팅 중복 제외");
            message.LogId++; message.Message = "/명령어"; store.CountMessage(message);
            Check(store.Rows["monthly:2026-10:10:" + user.UserId].Count == 1, "명령어 제외");
            message.Message = "채팅"; message.ChatId = 11; store.CountMessage(message);
            Check(store.Rows["monthly:2026-10:11:" + user.UserId].Count == 1, "방별 분리");
            store.AddPopularity(10, user.UserId, 200, 10, "이름", boundary);
            store.AddPopularity(10, user.UserId, 201, -3, "새이름", boundary);
            Check(!store.AddPopularity(10, user.UserId, 201, -3, "새이름", boundary), "인기도 중복 방지");
            Check(store.Ranking(10, "2026-10", true).Contains("7점"), "월간 인기도 순증감");
            Check(store.Ranking(10, "2026-09", true).Contains("없습니다"), "월별 인기도 분리");
            var history = new List<RoomEventRecord> { new RoomEventRecord { ChatId = 10, UserId = user.UserId, LogId = 300, Kind = "leave", Nickname = "옛이름", OccurredAt = before } };
            message.ChatId = 10; message.LogId = 301; message.RoomEvent = ChatRoomEvent.Join;
            Check(store.Reentry(message, history), "같은 ID 재입장 알림");
            Check(!store.Reentry(message, history), "재입장 알림 중복 방지");
            message.AuthorId = 22; Check(!store.Reentry(message, history), "다른 사용자 오인 금지");
            store = Reload(store); message.AuthorId = user.UserId;
            Check(!store.Reentry(message, history), "재시작 후 알림 중복 방지");
            Check(store.Rows.Values.Single(r => r.Kind == "reentry").UserId == user.UserId, "긴 ID 저장 왕복");
            Console.WriteLine("PASS: " + checks + " operations checks (no real resets, writes or sends)"); return 0;
        }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
}
