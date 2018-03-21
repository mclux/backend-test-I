using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace DevCenterBot
{
    public class GoogleSheetHelper
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Dev Center Google Sheets";

        public void writeOutputToSheet(List<ExcelUserVM> records)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                //credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1vv4ZpI4MmSdumTI0V3I07QTUtitUM72DiVPpBcZ1nr8";
            String range = "Sheet1!A2:B2";
            //SpreadsheetsResource.ValuesResource.GetRequest request =
            //        service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1vv4ZpI4MmSdumTI0V3I07QTUtitUM72DiVPpBcZ1nr8/edit
            //ValueRange response = request.Execute();
            //IList<IList<Object>> values = response.Values;
            //if (values != null && values.Count > 0)
            //{
            //    Console.WriteLine("Name, Major");
            //    foreach (var row in values)
            //    {
            //        // Print columns A and E, which correspond to indices 0 and 4.
            //        Console.WriteLine("{0}, {1}", row[0], row[1]);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No data found.");
            //}
            Console.WriteLine("Writing data to google sheet...");
            SpreadsheetsResource.ValuesResource.AppendRequest request;
            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            IList<Object> obj;

            foreach (var record in records)
            {
                obj = new List<Object>();
                obj.Add(record.Name);
                obj.Add(record.FollowerCount);
                objNewRecords.Add(obj);
            }            

            request = service.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, spreadsheetId, range);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();
            Console.WriteLine("-------- Done! ----------");
        }
    }
}
