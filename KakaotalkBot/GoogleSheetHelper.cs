﻿using System.Collections.Generic;
using System.IO;
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

        public GoogleSheetHelper(string applicationName, string sheetId)
        {
            this.applicationName = applicationName;
            this.sheetId = sheetId;
        }

        private SheetsService GetSheetsService()
        {
            GoogleCredential credential;
            using (var stream = new FileStream(CredentialFile, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
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

        public List<List<string>> ReadAllFromSheet(string sheetName)
        {
            var service = GetSheetsService();
            var range = $"{sheetName}"; // 전체 시트 범위

            var request = service.Spreadsheets.Values.Get(sheetId, range);
            var response = request.Execute();
            var values = response.Values;

            List<List<string>> result = new List<List<string>>();

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
