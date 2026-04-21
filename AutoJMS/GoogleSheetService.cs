using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;

namespace AutoJMS
{
    public static class GoogleSheetService
    {
        public const string DATA_SPREADSHEET_ID = "1EOuwlnynaMdYrFyR8ymeX6V-vO-FF6g2Ynb1xEykUpQ";
        public const string LICENSE_SPREADSHEET_ID = "1nx2VoXnAU3h8GRPxXkZ4c9Ev8jwHg3Iyor5o6wvsLNY";

        private const string CREDENTIAL_PATH = "service_account.json";
        private static SheetsService Service;


        public static void InitService()
        {
            if (Service != null) return;

            if (!File.Exists(CREDENTIAL_PATH))
                throw new FileNotFoundException($"Không tìm thấy tệp xác thực tại: {CREDENTIAL_PATH}");

            using (var stream = new FileStream(CREDENTIAL_PATH, FileMode.Open, FileAccess.Read))
            {
                var credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
                Service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "AutoJMS"
                });
            }
        }

        public static bool IsConfigValid()
        {
            return File.Exists(CREDENTIAL_PATH) && !string.IsNullOrWhiteSpace(DATA_SPREADSHEET_ID);
        }
        public static List<string> ReadColumnBySpreadsheetId(string spreadsheetId, string sheetName, int columnIndex)
        {
            try
            {
                if (Service == null) InitService();

                string colLetter = GetColumnLetter(columnIndex);
                string range = $"{sheetName}!{colLetter}2:{colLetter}";

                var request = Service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                var values = response.Values;

                List<string> result = new List<string>();
                if (values != null)
                {
                    foreach (var row in values)
                    {
                        if (row.Count > 0 && !string.IsNullOrWhiteSpace(row[0]?.ToString()))
                            result.Add(row[0].ToString().Trim());
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GoogleSheetService] Error: {ex.Message}");
                return new List<string>();

            }
        }

        // --- Bổ sung vào class GoogleSheetService ---
        public static void ClearSheet(string spreadsheetId, string range)
        {
            if (Service == null) InitService();
            var clearRequest = Service.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, range);
            clearRequest.Execute();
        }
        // 1. Ghi dữ liệu vào sheet BUMP
        public static void UpdateBumpSheet(IList<IList<object>> values, string spreadsheetId, string range)
        {
            if (Service == null) InitService();
            var valueRange = new ValueRange { Values = values };
            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }

        // 2. Đọc dữ liệu từ sheet COMMAND_QUEUE
        public static IList<IList<object>> ReadRange(string spreadsheetId, string range)
        {
            if (Service == null) InitService();
            var request = Service.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = request.Execute();
            return response.Values;
        }

        // 3. Cập nhật trạng thái từng ô trong COMMAND_QUEUE (VD: PENDING -> DONE)
        public static void UpdateCell(string spreadsheetId, string range, string value)
        {
            if (Service == null) InitService();
            var oblist = new List<object>() { value };
            var valueRange = new ValueRange { Values = new List<IList<object>> { oblist } };
            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }
        public static List<string> ReadColumn(string sheetName, int colIndex)
        {
            return ReadColumnBySpreadsheetId(DATA_SPREADSHEET_ID, sheetName, colIndex);
        }

        private static string GetColumnLetter(int col)
        {
            int dividend = col;
            string columnName = String.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
}