using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KakaotalkBot;

internal static class VerifyIdentityMigration
{
    static void Main(string[] args)
    {
        try { Verify(args); } catch (Exception e) { Console.Error.WriteLine(e.GetBaseException().Message); Environment.ExitCode = 1; }
    }
    static void Verify(string[] args)
    {
        string directory = args[0];
        var config = JObject.Parse(File.ReadAllText("BotIdentity.json"));
        string target = (string)config["SpreadsheetId"];
        var helper = new GoogleSheetHelper("IdentityVerification", target);
        var nativeUsers = helper.ReadUserTable();
        var known = new HashSet<string>(nativeUsers.Select(r => r[7]));
        var assembly = typeof(GoogleSheetHelper).Assembly;
        var method = typeof(GoogleSheetHelper).GetMethod("ReadActivityTable", BindingFlags.Instance | BindingFlags.NonPublic);
        var definitions = new[] {
            new[] { "DB_RoomUsers", "UserActivityStore", "RoomHeaders" }, new[] { "DB_RoomEvents", "UserActivityStore", "EventHeaders" },
            new[] { "DB_NicknameHistory", "UserActivityStore", "NicknameHeaders" }, new[] { "DB_RoomOperators", "RoomOperatorStore", "Headers" },
            new[] { "DB_Operations", "OperationsStore", "Headers" }, new[] { "DB_RoomTitles", "RoomTitleStore", "Headers" }, new[] { "DB_DailyQuests", "DailyQuestStore", "Headers" }
        };
        var decoded = new Dictionary<string, List<List<string>>> { { "DB_UserId", nativeUsers } };
        foreach (var definition in definitions)
        {
            var type = assembly.GetType("KakaotalkBot." + definition[1]);
            var headers = (string[])type.GetField(definition[2], BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
            decoded[definition[0]] = (List<List<string>>)method.Invoke(helper, new object[] { definition[0], headers });
        }
        // 실제 봇의 저장 경로로 기존 사용자 값을 한 번 저장합니다. 포인트나 사용자 수는 바꾸지 않습니다.
        if (args.Contains("--write-check")) helper.WriteUserTable(nativeUsers.Select(r => r.Cast<object>().ToList()).ToList());
        using (var api = new SheetsService(new BaseClientService.Initializer { ApplicationName = "IdentityVerification", Serializer = new Google.Apis.Json.NewtonsoftJsonSerializer(new JsonSerializerSettings { DateParseHandling = DateParseHandling.None }), HttpClientInitializer = GoogleCredential.FromFile("credentials.json").CreateScoped(SheetsService.Scope.Spreadsheets) }))
        {
            var entries = JsonConvert.DeserializeObject<JArray>(File.ReadAllText(Path.Combine(directory, "sanitized.json")), new JsonSerializerSettings { DateParseHandling = DateParseHandling.None });
            foreach (var entry in entries)
            {
                var sheet = entry["Sheet"].ToObject<Sheet>(); var title = sheet.Properties.Title;
                var expected = entry["Rows"].ToObject<List<List<object>>>();
                int width = expected.Count == 0 ? 1 : expected.Max(r => r.Count);
                var request = api.Spreadsheets.Values.Get(target, "'" + title.Replace("'", "''") + "'!A1:" + (char)('A' + width - 1) + Math.Max(1, expected.Count));
                request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA;
                var actual = request.Execute().Values ?? new List<IList<object>>();
                if (actual.Count != expected.Count) throw new Exception(title + " 행 수 불일치");
                for (int row = 0; row < expected.Count; row++) for (int col = 0; col < expected[row].Count; col++)
                {
                    string a = col < actual[row].Count ? Convert.ToString(actual[row][col]) : "", e = Convert.ToString(expected[row][col]);
                    if (a != e && !(a.StartsWith("{") && e.StartsWith("{") && JToken.DeepEquals(JToken.Parse(a), JToken.Parse(e)))) throw new Exception(title + " 값 불일치: " + (row + 1) + "," + (col + 1));
                    if (System.Text.RegularExpressions.Regex.Matches(a, @"(?<![A-Za-z0-9_])[0-9]+(?![A-Za-z0-9_])").Cast<System.Text.RegularExpressions.Match>().Any(m => known.Contains(m.Value))) throw new Exception(title + " 실제 ID 발견");
                }
                if (decoded.ContainsKey(title))
                {
                    var source = JsonConvert.DeserializeObject<Sheet>(File.ReadAllText(Path.Combine(directory, "sourcecompact", entry["Sheet"]["properties"]["sheetId"] + ".json")));
                    var sourceRows = source.Data[0].RowData.Skip(1).Where(r => r.Values != null && r.Values.Any(c => c.UserEnteredValue != null)).ToList();
                    if (sourceRows.Count != decoded[title].Count) throw new Exception(title + " 복원 행 수 불일치");
                    for (int r = 0; r < sourceRows.Count; r++) for (int c = 0; c < decoded[title][r].Count; c++)
                    {
                        var v = c < sourceRows[r].Values.Count ? sourceRows[r].Values[c].UserEnteredValue : null;
                        string original = v == null ? "" : v.StringValue ?? Convert.ToString((object)v.NumberValue ?? v.BoolValue);
                        string restored = decoded[title][r][c];
                        if (original != restored && !(original.StartsWith("{") && restored.StartsWith("{") && JToken.DeepEquals(JToken.Parse(original), JToken.Parse(restored)))) throw new Exception(title + " 원본 복원 불일치");
                    }
                }
                if (title.Contains("퀴즈") || title == "경제" || title == "시트13")
                {
                    request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
                    var evaluated = request.Execute().Values;
                    if (evaluated.Any(r => r.Any(v => new[] { "#NAME?", "#REF!", "#ERROR!", "#VALUE!" }.Contains(Convert.ToString(v))))) throw new Exception(title + " 수식 계산 오류");
                }
                Console.WriteLine(title + ": " + expected.Count + " rows matched, no raw user IDs");
            }
        }
        Console.WriteLine("Verified native users: " + nativeUsers.Count);
    }
}
