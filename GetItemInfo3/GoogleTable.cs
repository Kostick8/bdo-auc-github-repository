using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Runtime.Remoting.Messaging;

namespace GetItemInfo3
{
    internal class GoogleTable
    {
        public static SheetsService Service { get; set; }
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets, SheetsService.Scope.Drive };

        public static SheetsService AuthorizeGoogleApp(string Project_id)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = Project_id,
            });

            return service;
        }

        public static void SetValue(SheetsService service, string spreadsheetId, List<IList<object>> list, string StartCell, string MajorDimension)
        {
            string range = StartCell;
            ValueRange valueRange = new ValueRange
            {
                MajorDimension = MajorDimension,  //"ROWS";//COLUMNS
                Values = list
            };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            _ = update.Execute();
        }

        public static IList<IList<object>> GetSheet(SheetsService service, string spreadsheetId,  string range)
        { 
            SpreadsheetsResource.ValuesResource.GetRequest getRequest = service.Spreadsheets.Values.Get(spreadsheetId, range);
            getRequest.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
            var range1 = getRequest.Execute();

            return range1.Values;
        }

        public static void ClearSheet(SheetsService service, string spreadsheetId, string range)
        {
            service.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, range).Execute();
        }
    }
}
