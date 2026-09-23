using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace KakaotalkBot
{
    public class GoogleSheetHelper
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string CredentialFile = "credentials.json";

        private string applicationName = string.Empty;
        private string sheetName = string.Empty;
        private string sheetId = string.Empty;

        private object lockObject = new object();

        private SheetsService service;

        private static readonly string[] UserHeaders =
            { "nickname", "take_attendance", "attendance_at", "point", "popularity", "password", "contribution", "user_id", "leave_count", "kick_count", "experience", "level" };

        // 사용자 DB 읽기 실패를 빈 목록으로 취급하면 이후 저장 시 기존 데이터가 손상될 수 있다.
        public List<List<string>> ReadUserTable()
        {
            lock (lockObject)
            {
                var request = GetSheetsService().Spreadsheets.Values.Get(sheetId, "'" + Database.UserSheetName + "'!A1:L");
                request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
                return ParseUserRows(request.Execute().Values);
            }
        }

        internal static List<List<string>> ParseUserRows(IList<IList<object>> values)
        {
            if (values == null || values.Count == 0 || (values[0].Count != 10 && values[0].Count != UserHeaders.Length) ||
                !values[0].Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).SequenceEqual(UserHeaders.Take(values[0].Count)))
                throw new FormatException("DB_UserId의 사용자 헤더가 올바르지 않습니다. 기존 DB 탭은 사용하지 않습니다.");
            var result = new List<List<string>>();
            var ids = new HashSet<long>();
            for (int index = 1; index < values.Count; index++)
            {
                var row = values[index];
                if (row.All(value => string.IsNullOrEmpty(Convert.ToString(value, CultureInfo.InvariantCulture)))) continue;
                // 숫자 셀은 이미 반올림되었을 수 있으므로 문자열로 명시한 ID만 허용한다.
                if (row.Count < 8 || row.Count > UserHeaders.Length || !(row[7] is string)) throw new FormatException("DB_UserId " + (index + 1) + "행의 user_id를 텍스트로 저장해 주세요.");
                var cells = row.Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).ToList();
                var user = User.ToUser(cells);
                if (!ids.Add(user.UserId)) throw new FormatException("DB_UserId에 중복된 user_id가 있습니다: " + user.UserId);
                result.Add(cells);
            }
            return result;
        }

        public void WriteUserTable(List<List<object>> users)
        { WriteUserData(users, null); }

        internal void WriteUserData(List<List<object>> users, UserActivityStore activity, RoomOperatorStore operators = null, OperationsStore operations = null, RoomTitleStore titles = null)
        {
            lock (lockObject)
            {
                // 헤더와 데이터 형식을 먼저 검증하고, 실패는 호출자에게 전달한다.
                ParseUserRows(new List<IList<object>> { UserHeaders.Cast<object>().ToList() }.Concat(users).ToList());
                var api = GetSheetsService();
                ParseUserRows(api.Spreadsheets.Values.Get(sheetId, "'" + Database.UserSheetName + "'!A1:L1").Execute().Values);
                var metadataRequest = api.Spreadsheets.Get(sheetId);
                metadataRequest.Fields = "sheets.properties";
                var metadata = metadataRequest.Execute();
                var ranges = new List<ValueRange>();
                var extensions = new List<Request>();
                Action<string, int, List<List<object>>, int> add = (title, first, rows, columns) =>
                {
                    if (rows.Count == 0) return;
                    var sheet = metadata.Sheets.Single(s => s.Properties.Title == title);
                    int required = first + rows.Count - 1;
                    int count = sheet.Properties.GridProperties.RowCount.Value;
                    if (required > count) extensions.Add(new Request { AppendDimension = new AppendDimensionRequest {
                        SheetId = sheet.Properties.SheetId, Dimension = "ROWS", Length = required - count } });
                    ranges.Add(new ValueRange { Range = "'" + title + "'!A" + first + ":" + (char)('A' + columns - 1) + required,
                        Values = rows.Cast<IList<object>>().ToList() });
                };
                add(Database.UserSheetName, 2, users, 12);
                if (operations != null) add("DB_Operations", 2, operations.ToRows(), 3);
                if (titles != null) add("DB_RoomTitles", 2, titles.ToRows(), 2);
                if (operators != null)
                    add("DB_RoomOperators", 2, operators.Records.Values.OrderBy(r => r.ChatId).ThenBy(r => r.UserId).Select(r => r.ToRow()).ToList(), 10);
                if (activity != null)
                {
                    add("DB_RoomUsers", 2, activity.Rooms.Values.OrderBy(r => r.ChatId).ThenBy(r => r.UserId).Select(r => r.ToRow()).ToList(), 9);
                    // 확정된 행 번호에 쓰므로 응답이 유실되어 재시도해도 이력이 중복되지 않습니다.
                    add("DB_RoomEvents", activity.SavedEvents + 2, activity.Events.Skip(activity.SavedEvents).Select(r => r.ToRow()).ToList(), 7);
                    add("DB_NicknameHistory", activity.SavedNicknames + 2, activity.Nicknames.Skip(activity.SavedNicknames).Select(r => r.ToRow()).ToList(), 8);
                }
                if (extensions.Count > 0) api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = extensions }, sheetId).Execute();
                // 사용자 상태와 새 이력을 한 요청에 저장합니다. ID와 경험치는 문자열 + RAW로 정밀도를 유지합니다.
                if (ranges.Count > 0) api.Spreadsheets.Values.BatchUpdate(new BatchUpdateValuesRequest { ValueInputOption = "RAW", Data = ranges }, sheetId).Execute();
                if (titles != null) titles.Dirty = false;
                if (activity != null) activity.MarkSaved();
                if (operators != null) operators.Dirty = false;
                if (operations != null) { operations.Dirty = false; operations.ResetPending = false; }
            }
        }

        internal void EnsureActivitySchema()
        {
            lock (lockObject)
            {
                var api = GetSheetsService();
                var query = api.Spreadsheets.Get(sheetId); query.Fields = "sheets.properties";
                var sheets = query.Execute().Sheets;
                var user = sheets.Single(s => s.Properties.Title == Database.UserSheetName);
                var header = api.Spreadsheets.Values.Get(sheetId, "'DB_UserId'!A1:J1").Execute().Values;
                ParseUserRows(header);
                var requests = new List<Request>();
                if (user.Properties.GridProperties.ColumnCount < 12) requests.Add(new Request { AppendDimension = new AppendDimensionRequest {
                    SheetId = user.Properties.SheetId, Dimension = "COLUMNS", Length = 12 - user.Properties.GridProperties.ColumnCount } });
                var definitions = new Dictionary<string, string[]> { { "DB_RoomUsers", UserActivityStore.RoomHeaders },
                    { "DB_RoomEvents", UserActivityStore.EventHeaders }, { "DB_NicknameHistory", UserActivityStore.NicknameHeaders },
                    { "DB_RoomOperators", RoomOperatorStore.Headers }, { "DB_Operations", OperationsStore.Headers }, { "DB_RoomTitles", RoomTitleStore.Headers } };
                var sheetIds = new HashSet<int>(sheets.Select(s => s.Properties.SheetId.Value));
                int newSheetId = 1;
                foreach (var definition in definitions)
                    if (!sheets.Any(s => s.Properties.Title == definition.Key))
                    {
                        while (sheetIds.Contains(newSheetId)) newSheetId++;
                        int id = newSheetId++; sheetIds.Add(id);
                        requests.Add(new Request { AddSheet = new AddSheetRequest { Properties = new SheetProperties {
                            SheetId = id, Title = definition.Key, GridProperties = new GridProperties {
                                RowCount = 1000, ColumnCount = definition.Value.Length, FrozenRowCount = 1 } } } });
                        // 생성과 헤더를 한 요청에 묶어 중간 실패로 빈 탭만 남지 않게 합니다.
                        requests.Add(new Request { UpdateCells = new UpdateCellsRequest { Start = new GridCoordinate { SheetId = id, RowIndex = 0, ColumnIndex = 0 },
                            Rows = new List<RowData> { new RowData { Values = definition.Value.Select(h => new CellData {
                                UserEnteredValue = new ExtendedValue { StringValue = h } }).ToList() } }, Fields = "userEnteredValue" } });
                    }
                if (requests.Count > 0) api.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, sheetId).Execute();
                var extra = api.Spreadsheets.Values.Get(sheetId, "'DB_UserId'!K1:L1").Execute().Values;
                if (extra != null && extra.Count > 0 && !extra[0].Select(Convert.ToString).SequenceEqual(UserHeaders.Skip(10)))
                    throw new FormatException("DB_UserId K:L에 다른 컬럼이 있습니다.");
                var headers = new List<ValueRange>();
                if (extra == null || extra.Count == 0) headers.Add(new ValueRange { Range = "'DB_UserId'!K1:L1", Values = new List<IList<object>> { UserHeaders.Skip(10).Cast<object>().ToList() } });
                foreach (var definition in definitions)
                {
                    // 기존 탭은 빈 표로 덮어쓰지 않고 헤더를 검증합니다.
                    if (sheets.Any(s => s.Properties.Title == definition.Key)) ReadActivityTable(definition.Key, definition.Value);
                    else headers.Add(new ValueRange { Range = "'" + definition.Key + "'!A1", Values = new List<IList<object>> { definition.Value.Cast<object>().ToList() } });
                }
                if (headers.Count > 0) api.Spreadsheets.Values.BatchUpdate(new BatchUpdateValuesRequest { ValueInputOption = "RAW", Data = headers }, sheetId).Execute();
            }
        }

        internal List<List<string>> ReadActivityTable(string title, string[] headers)
        {
            var request = GetSheetsService().Spreadsheets.Values.Get(sheetId, "'" + title + "'!A1:" + (char)('A' + headers.Length - 1));
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
            var values = request.Execute().Values;
            if (values == null || values.Count == 0 || !values[0].Select(Convert.ToString).SequenceEqual(headers))
                throw new FormatException(title + " 헤더가 잘못되었습니다.");
            var result = new List<List<string>>();
            foreach (var row in values.Skip(1))
            {
                if (row.Count == 0 || row.All(v => string.IsNullOrEmpty(Convert.ToString(v))))
                    throw new FormatException(title + " 중간에 빈 행이 있습니다.");
                // ID 컬럼은 숫자 셀로 저장된 반올림 값을 허용하지 않습니다.
                int[] idColumns = title == "DB_RoomOperators" ? new[] { 0, 1, 8, 9 } : title == "DB_RoomUsers" ? new[] { 0, 1, 5, 6 } : title == "DB_RoomEvents" ? new[] { 1, 2, 5 } : new[] { 1, 2, 6 };
                foreach (int column in (title == "DB_Operations" || title == "DB_RoomTitles") ? new int[0] : idColumns)
                    if (row.Count <= column || !(row[column] is string)) throw new FormatException(title + "의 ID는 텍스트여야 합니다.");
                result.Add(row.Select(v => Convert.ToString(v, CultureInfo.InvariantCulture)).Concat(Enumerable.Repeat("", headers.Length - row.Count)).ToList());
            }
            return result;
        }

        // 두 표를 모두 읽은 뒤 교체하므로 통신 실패 시 기존 명령 목록이 유지됩니다.
        internal List<List<string>> ReadCommandTables()
        {
            lock (lockObject)
            {
                var request = GetSheetsService().Spreadsheets.Values.BatchGet(sheetId);
                request.Ranges = new[] { "'키워드'!A:B", "'어록'!A:B" };
                var tables = request.Execute().ValueRanges;
                if (tables == null || tables.Count != 2)
                    throw new InvalidDataException("키워드와 어록 시트를 모두 읽지 못했습니다.");
                var result = new List<List<string>>();
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var table in tables)
                {
                    if (table.Values == null) continue;
                    foreach (var row in table.Values)
                    {
                        if (row.Count == 0) continue;
                        string keyword = Convert.ToString(row[0], CultureInfo.InvariantCulture);
                        // 빈 행과 제목 행은 제외하고, 중복 이름은 키워드 시트를 우선합니다.
                        if (string.IsNullOrWhiteSpace(keyword) || !keyword.StartsWith("/", StringComparison.Ordinal) || !seen.Add(keyword)) continue;
                        result.Add(row.Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)).ToList());
                    }
                }
                return result;
            }
        }

        public GoogleSheetHelper(string applicationName, string sheetId)
        {
            this.applicationName = applicationName;
            this.sheetId = sheetId;
        }

        private SheetsService GetSheetsService()
        {
            if (service == null)
            {
                GoogleCredential credential;
                using (var stream = new FileStream(CredentialFile, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
                }

                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = applicationName,
                });
            }

            return service;
        }

        public void WriteToSheet(string sheetName, List<string> messages)
        {
            var service = GetSheetsService();

            var valueRange = new ValueRange();
            var values = new List<IList<object>>();

            foreach (var msg in messages)
            {
                values.Add(new List<object> { msg });
            }

            valueRange.Values = values;

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, sheetId, $"{sheetName}!A1");
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.Execute();
        }

        public void WriteToSheetAll(string sheetName, List<List<object>> messages)
        {
            lock (lockObject)
            {
                var service = GetSheetsService();

                var valueRange = new ValueRange();
                var values = new List<IList<object>>();

                foreach (var msg in messages)
                {
                    // msg (List<string>)를 object 리스트로 변환
                    values.Add(msg.Cast<object>().ToList());
                }

                valueRange.Values = values;

                var updateRequest = service.Spreadsheets.Values.Update(valueRange, sheetId, sheetName);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;


                try
                {
                    var response = updateRequest.Execute();

                }
                catch(Exception e)
                {
                    return;
                }
            }
            
        }


        public List<List<string>> ReadAllFromSheet(string sheetName)
        {
            List<List<string>> result = new List<List<string>>();
            IList<IList<object>> values = null;

            lock (lockObject)
            {
                var service = GetSheetsService();
                var range = $"{sheetName}"; // 전체 시트 범위
                ValueRange response = null;
                
                try
                {
                    var request = service.Spreadsheets.Values.Get(sheetId, range);
                    response = request.Execute();
                    values = response.Values;
                }
                catch (Exception e)
                {

                }
                finally
                {

                }
            }
            
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    var rowList = new List<string>();
                    foreach (var cell in row)
                    {
                        rowList.Add(cell.ToString());
                    }
                    result.Add(rowList);
                }
            }

            return result;
        }
    }
}
