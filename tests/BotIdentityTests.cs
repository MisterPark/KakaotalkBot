using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using KakaotalkBot;
internal static class BotIdentityTests
{
    static int count;
    static void Check(bool condition) { if (!condition) throw new Exception("검증 실패: " + (count + 1)); count++; }
    static void Reject(Action action) { bool rejected = false; try { action(); } catch { rejected = true; } Check(rejected); }
    static IList<IList<object>> Rows(params object[] row) { return new List<IList<object>> { row.ToList() }; }
    static void Main()
    {
        var config = new BotIdentityConfiguration { Key = Convert.ToBase64String(Enumerable.Range(0, 64).Select(i => (byte)i).ToArray()), SpreadsheetId = "test" };
        var a = new BotIdentity(config); var b = new BotIdentity(config);
        foreach (long id in new[] { 1L, 1234567890L, 9007199254740993L, long.MaxValue })
        { Check(a.Decode(a.Encode(id)) == id); Check(a.Encode(id) == b.Encode(id)); }
        Check(a.Encode(0) == "0" && a.Decode("0") == 0);
        Reject(() => a.Decode("1234567890"));
        string token = a.Encode(1234567890);
        foreach (string nickname in new[] { "bot_hello", "bot_literal_hello", token, "0001-01-01T00:00:00.0000000" })
        {
            var input = Rows(nickname, "", "0001-01-01T00:00:00.0000000", "1", "0", "pw", "0", "1234567890", "0", "0", "0", "1");
            var output = b.Transform("DB_UserId", a.Transform("DB_UserId", input, true, false), false, false);
            Check(Convert.ToString(output[0][0]) == nickname && Convert.ToString(output[0][2]) == Convert.ToString(input[0][2]));
        }
        Reject(() => a.Decode(token.Substring(0, token.Length - 3) + "AAA"));
        var wrong = new BotIdentity(new BotIdentityConfiguration { Key = Convert.ToBase64String(new byte[64]) });
        Reject(() => wrong.Decode(token));
        foreach (var example in new[] {
            new { Title = "DB_UserId", Cells = Rows("유저", "", "", "100", "2", "pw", "4", "1234567890", "1", "2", "300", "3") },
            new { Title = "DB_RoomUsers", Cells = Rows("1234567890", "1234567890", "유저", "true", "1", "1234567890", "7", "1", "1") },
            new { Title = "DB_RoomEvents", Cells = Rows("1234567890:1234567890:1234567890", "1234567890", "1234567890", "leave", "1", "1234567890", "유저") },
            new { Title = "DB_NicknameHistory", Cells = Rows("guid", "1234567890", "1234567890", "이전", "현재", "1", "1234567890", "chat") },
            new { Title = "DB_RoomOperators", Cells = Rows("1234567890", "1234567890", "유저", "4", "0", "true", "true", "1", "1234567890", "1234567890") },
            new { Title = "DB_Operations", Cells = Rows("memo", "memo:12:13", "{\"Key\":\"memo:12:13\",\"Kind\":\"memo\",\"UserId\":1234567890,\"AuthorId\":9007199254740993,\"ChatId\":1234567890,\"Text\":\"대상 1234567890\"}") },
            new { Title = "DB_Operations", Cells = Rows("monthly", "monthly:2026-09:12:1234567890", "{\"Key\":\"monthly:2026-09:12:1234567890\",\"Kind\":\"monthly\",\"UserId\":1234567890,\"AuthorId\":0,\"ChatId\":12}") },
            new { Title = "DB_RoomTitles", Cells = Rows("monthly:2026-09:12:chat:1234567890", "{\"Key\":\"monthly:2026-09:12:chat:1234567890\",\"UserId\":1234567890,\"AuthorId\":0,\"RevokedBy\":9007199254740993}") },
            new { Title = "DB_DailyQuests", Cells = Rows("2026-09-25:1234567890", "{\"Key\":\"2026-09-25:1234567890\",\"UserId\":1234567890,\"ChatCursors\":{\"1234567890\":1234567890},\"Rewards\":{\"chat\":{\"Key\":\"2026-09-25:1234567890:chat\",\"Points\":10}}}") }
        })
        {
            var encoded = a.Transform(example.Title, example.Cells, true, false);
            var decoded = b.Transform(example.Title, encoded, false, false);
            for (int i = 0; i < example.Cells[0].Count; i++)
            {
                string expected = Convert.ToString(example.Cells[0][i]), actual = Convert.ToString(decoded[0][i]);
                Check(expected.StartsWith("{") ? JToken.DeepEquals(JToken.Parse(expected), JToken.Parse(actual)) : expected == actual);
            }
            if (example.Title == "DB_RoomUsers") Check(Convert.ToString(encoded[0][0]) == "1234567890" && Convert.ToString(encoded[0][1]) == token);
        }
        Reject(() => a.Transform("DB_UserId", Rows("nickname", "", "", "", "", "", "", "user_id"), false, true));
        Console.WriteLine("BotIdentity: " + count + " checks passed");
    }
}
