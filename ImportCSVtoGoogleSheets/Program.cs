using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Nortal.Utilities.Csv;
using System;
using System.IO;
using System.Threading;

namespace ImportCSVtoGoogleSheets
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Import CSV to Google Sheets";

        static void Main(string[] args)
        {
            if(args.Length != 2 || string.IsNullOrEmpty(args[0]) || string.IsNullOrEmpty(args[1]))
            {
                Console.WriteLine("Import CSV to Google Sheets v0.1 by Maxim Rubis");
                Console.WriteLine();
                Console.WriteLine("Usage: ImportCSVtoGoogleSheets.exe <Sheet ID> <File to import>");
                return;
            }

            try
            {
                UserCredential credential;

                using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                string spreadsheetId = args[0];
                service.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, "A:Z").Execute();

                var csv = File.ReadAllText("Returnship - Number of Requests.csv");
                String[][] data = CsvParser.Parse(csv);



                var range = new ValueRange();
                range.MajorDimension = "ROWS";
                range.Values = data;

                var update = service.Spreadsheets.Values.Update(range, spreadsheetId, "a:z");
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                update.Execute();

            }
            catch(Exception e)
            {
                Console.WriteLine("Error occured: " + e.Message + (e.InnerException != null ? " -> " + e.InnerException.Message : ""));
            }
        }
    }
}
