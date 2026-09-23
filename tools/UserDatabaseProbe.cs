using System;
using KakaotalkBot;
using System.Collections;
using System.Reflection;

// 봇의 서비스 계정으로 새 사용자 시트의 스키마와 ID 형식만 확인한다. 쓰기나 채팅 전송은 하지 않는다.
internal static class UserDatabaseProbe
{
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length != 1) throw new ArgumentException("스프레드시트 ID가 필요합니다.");
            var helper = new GoogleSheetHelper("KakaoUserIdentityProbe", args[0]);
            var rows = helper.ReadUserTable();
            Console.WriteLine("PASS: " + Database.UserSheetName + " 스키마·ID 검사, 사용자 수=" + rows.Count);
            Type store = typeof(Database).Assembly.GetType("KakaotalkBot.UserActivityStore");
            MethodInfo read = typeof(GoogleSheetHelper).GetMethod("ReadActivityTable", BindingFlags.Instance | BindingFlags.NonPublic);
            foreach (string name in new[] { "Room", "Event", "Nickname" })
            {
                string title = name == "Room" ? "DB_RoomUsers" : name == "Event" ? "DB_RoomEvents" : "DB_NicknameHistory";
                object headers = store.GetField(name + "Headers", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                var values = (IList)read.Invoke(helper, new[] { (object)title, headers });
                Console.WriteLine("PASS: " + title + " 스키마·ID 검사, 행 수=" + values.Count);
            }
            Type operators = typeof(Database).Assembly.GetType("KakaotalkBot.RoomOperatorStore");
            var operatorHeaders = operators.GetField("Headers", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
            var operatorRows = (IList)read.Invoke(helper, new[] { (object)"DB_RoomOperators", operatorHeaders });
            Console.WriteLine("PASS: DB_RoomOperators 스키마·ID 검사, 행 수=" + operatorRows.Count);
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error.Message); return 1; }
    }
}
