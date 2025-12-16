using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
