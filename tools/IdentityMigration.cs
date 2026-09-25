using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KakaotalkBot;

// 원본은 읽기만 합니다. 실행 중인 봇이 없어야 일관된 스냅샷을 얻을 수 있습니다.
internal static class IdentityMigration
{
    static void Main(string[] args) { try { Run(args); } catch (Exception e) { Console.Error.WriteLine(e.GetType().Name + ": " + e.Message); Environment.ExitCode = 1; } }
    static void Run(string[] args)
    {
        JsonConvert.DefaultSettings = () => new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore, DateParseHandling = DateParseHandling.None };
        string mode = args[0], directory = args[1];
        Directory.CreateDirectory(directory);
        var credential = GoogleCredential.FromFile("credentials.json").CreateScoped(SheetsService.Scope.Spreadsheets);
        using (var api = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential, ApplicationName = "BotIdentityMigration" }))
        {
            if (mode == "snapshot")
            {
                string source = args[2];
                var query = api.Spreadsheets.Get(source); query.Fields = "properties,sheets.properties,namedRanges";
                var metadata = query.Execute();
                File.WriteAllText(Path.Combine(directory, "metadata.json"), JsonConvert.SerializeObject(metadata));
                foreach (var sheet in metadata.Sheets)
                {
                    string title = sheet.Properties.Title;
                    if (args.Length > 3 && !title.StartsWith("DB_", StringComparison.Ordinal)) continue;
                    var request = api.Spreadsheets.Get(source);
                    request.Ranges = new[] { "'" + title.Replace("'", "''") + "'!A1:" + (char)('A' + sheet.Properties.GridProperties.ColumnCount.Value - 1) + sheet.Properties.GridProperties.RowCount };
                    request.Fields = "sheets(properties,data(startRow,startColumn,rowData(values(userEnteredValue,userEnteredFormat,dataValidation,note,textFormatRuns)),rowMetadata,columnMetadata),merges,conditionalFormats,basicFilter)";
                    var data = request.Execute().Sheets.Single();
                    File.WriteAllText(Path.Combine(directory, sheet.Properties.SheetId + ".json"), JsonConvert.SerializeObject(data));
                    Console.WriteLine(title + ": snapshot complete");
                }
                return;
            }
            var identity = new BotIdentity(JsonConvert.DeserializeObject<BotIdentityConfiguration>(File.ReadAllText(args[2])));
            if (mode == "repair-user-values")
            {
                if (identity.SpreadsheetId == "1ECWcE98IMRm9PItPpK6rEC7dmmxbuT3Jtu3aFnaeccQ") throw new Exception("원본에는 쓸 수 없습니다.");
                var data = JsonConvert.DeserializeObject<JArray>(File.ReadAllText(Path.Combine(directory, "sanitized.json")), new JsonSerializerSettings { DateParseHandling = DateParseHandling.None });
                var users = data.Single(e => (string)e["Sheet"]["properties"]["title"] == "DB_UserId")["Rows"].ToObject<List<IList<object>>>();
                var update = api.Spreadsheets.Values.Update(new ValueRange { Values = users }, identity.SpreadsheetId, "'DB_UserId'!A1:L" + users.Count);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW; update.Execute();
                Console.WriteLine("User values restored from sanitized snapshot"); return;
            }
            if (mode == "restore") { Restore(api, directory, identity.SpreadsheetId, args.Length > 3 ? args[3] : null); return; }
            var sheets = Directory.GetFiles(directory, "*.json").Where(f => Path.GetFileName(f) != "metadata.json").Select(f => JsonConvert.DeserializeObject<Sheet>(File.ReadAllText(f))).ToList();
            var ordered = sheets.OrderBy(s => s.Properties.Title == "DB_UserId" ? -1 : s.Properties.Index).ToList();
            var output = new List<object>();
            foreach (var sheet in ordered)
            {
                var grid = sheet.Data.FirstOrDefault();
                var rows = grid == null || grid.RowData == null ? new List<IList<object>>() : grid.RowData.Select(r => (IList<object>)(r.Values ?? new List<CellData>()).Select(c => c.UserEnteredValue == null ? (object)"" : c.UserEnteredValue.StringValue ?? (object)c.UserEnteredValue.NumberValue ?? c.UserEnteredValue.BoolValue ?? (object)c.UserEnteredValue.FormulaValue ?? "").ToList()).ToList();
                // 서식만 있는 꼬리 행을 값 변환 대상에서 제외합니다.
                while (rows.Count > 0 && rows[rows.Count - 1].All(v => Convert.ToString(v) == "")) rows.RemoveAt(rows.Count - 1);
                var converted = identity.Transform(sheet.Properties.Title, rows, true, true);
                var restored = identity.Transform(sheet.Properties.Title, converted, false, true);
                if (sheet.Properties.Title.StartsWith("DB_") && rows.Count != restored.Count) throw new Exception("행 수 불일치");
                for (int r = 0; r < rows.Count; r++) for (int c = 0; c < rows[r].Count; c++)
                {
                    string before = Convert.ToString(rows[r][c]), after = Convert.ToString(restored[r][c]);
                    if (sheet.Properties.Title.StartsWith("DB_") && before != after && !(before.StartsWith("{") && after.StartsWith("{") && JToken.DeepEquals(JToken.Parse(before), JToken.Parse(after)))) throw new Exception(sheet.Properties.Title + " roundtrip mismatch at " + (r + 1) + ":" + (c + 1));
                    if (Convert.ToString(converted[r][c]) != before) grid.RowData[r].Values[c].UserEnteredValue = new ExtendedValue { StringValue = Convert.ToString(converted[r][c]) };
                }
                // 셀 메모, 서식 링크와 유효성 검사 텍스트에도 기존 ID가 남지 않게 변환합니다.
                var token = JsonConvert.DeserializeObject<JObject>(JsonConvert.SerializeObject(sheet), new JsonSerializerSettings { DateParseHandling = DateParseHandling.None });
                foreach (var value in token.Descendants().OfType<JValue>().Where(v => v.Type == JTokenType.String).ToList())
                    value.Value = identity.Scrub((string)value.Value);
                sheet.Data = JsonConvert.DeserializeObject<Sheet>(token.ToString()).Data;
                output.Add(new { Sheet = sheet, Rows = converted });
                Console.WriteLine(sheet.Properties.Title + ": " + rows.Count + " rows verified");
            }
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(directory), "sanitized.json"), JsonConvert.SerializeObject(output));
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(directory), "marker.json"), JsonConvert.SerializeObject(new[] { new[] { "version", "key_check" }, new[] { "1", identity.KeyCheck } }));
        }
    }
    static void Restore(SheetsService api, string directory, string target, string resume)
    {
        if (target == "1ECWcE98IMRm9PItPpK6rEC7dmmxbuT3Jtu3aFnaeccQ") throw new Exception("원본에는 쓸 수 없습니다.");
        var q = api.Spreadsheets.Get(target); q.Fields = "sheets.properties";
        var metadata = q.Execute();
        var source = JsonConvert.DeserializeObject<Spreadsheet>(File.ReadAllText(Path.Combine(directory, "sourcecompact", "metadata.json")));
        api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { new Request { UpdateSpreadsheetProperties = new UpdateSpreadsheetPropertiesRequest {
            Properties = new SpreadsheetProperties { Locale = source.Properties.Locale, TimeZone = source.Properties.TimeZone, DefaultFormat = source.Properties.DefaultFormat }, Fields = "locale,timeZone,defaultFormat" } } } }, target).Execute();
        var entries = JsonConvert.DeserializeObject<JArray>(File.ReadAllText(Path.Combine(directory, "sanitized.json")), new JsonSerializerSettings { DateParseHandling = DateParseHandling.None });
        var idMap = new Dictionary<int, int>();
        foreach (var entry in entries)
        {
            var sheet = entry["Sheet"].ToObject<Sheet>();
            int id = metadata.Sheets.Single(s => s.Properties.Title == sheet.Properties.Title).Properties.SheetId.Value;
            idMap.Add(sheet.Properties.SheetId.Value, id);
            if (resume != null && sheet.Properties.Title != resume) continue;
            resume = null;
            System.Threading.Thread.Sleep(1500);
            var properties = sheet.Properties; properties.SheetId = id;
            api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { new Request { UpdateSheetProperties = new UpdateSheetPropertiesRequest {
                Properties = properties, Fields = "gridProperties,hidden,rightToLeft,tabColor" } } } }, target).Execute();
            var grid = sheet.Data.FirstOrDefault();
            if (grid != null && grid.RowData != null)
                for (int row = 0; row < grid.RowData.Count; row += 1000)
                    api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { new Request { UpdateCells = new UpdateCellsRequest {
                        Start = new GridCoordinate { SheetId = id, RowIndex = row, ColumnIndex = 0 }, Rows = grid.RowData.Skip(row).Take(1000).ToList(),
                        Fields = "userEnteredValue,userEnteredFormat,dataValidation,note,textFormatRuns" } } } }, target).Execute();
            var requests = new List<Request>();
            if (grid != null) foreach (var dimension in new[] { new { Name = "ROWS", Values = grid.RowMetadata }, new { Name = "COLUMNS", Values = grid.ColumnMetadata } })
            {
                if (dimension.Values == null) continue;
                for (int start = 0; start < dimension.Values.Count; )
                {
                    var p = dimension.Values[start]; int end = start + 1;
                    while (end < dimension.Values.Count && dimension.Values[end].PixelSize == p.PixelSize && dimension.Values[end].HiddenByUser == p.HiddenByUser) end++;
                    if (p.PixelSize.HasValue || p.HiddenByUser.HasValue) requests.Add(new Request { UpdateDimensionProperties = new UpdateDimensionPropertiesRequest {
                        Range = new DimensionRange { SheetId = id, Dimension = dimension.Name, StartIndex = start, EndIndex = end },
                        Properties = new DimensionProperties { PixelSize = p.PixelSize, HiddenByUser = p.HiddenByUser }, Fields = "pixelSize,hiddenByUser" } });
                    start = end;
                }
            }
            if (sheet.Merges != null) foreach (var range in sheet.Merges) { range.SheetId = id; requests.Add(new Request { MergeCells = new MergeCellsRequest { Range = range, MergeType = "MERGE_ALL" } }); }
            if (sheet.BasicFilter != null) { sheet.BasicFilter.Range.SheetId = id; requests.Add(new Request { SetBasicFilter = new SetBasicFilterRequest { Filter = sheet.BasicFilter } }); }
            if (sheet.ConditionalFormats != null) for (int i = 0; i < sheet.ConditionalFormats.Count; i++)
            { var rule = sheet.ConditionalFormats[i]; foreach (var range in rule.Ranges) range.SheetId = id; requests.Add(new Request { AddConditionalFormatRule = new AddConditionalFormatRuleRequest { Rule = rule, Index = i } }); }
            if (requests.Count > 0) api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, target).Execute();
            Console.WriteLine(properties.Title + ": native data and formatting restored");
        }
        if (source.NamedRanges != null && source.NamedRanges.Count > 0)
            api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = source.NamedRanges.Select(n => { n.NamedRangeId = null; n.Range.SheetId = idMap[n.Range.SheetId.Value]; return new Request { AddNamedRange = new AddNamedRangeRequest { NamedRange = n } }; }).ToList() }, target).Execute();
    }
}
