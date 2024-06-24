
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using System.Text.Json.Serialization;
using System.Text;
using System.Linq;

namespace StubAPICall
{

    public class dt { public string sym { get; set; } public float quote { get; set; } }


    static class Program
    {

        private static readonly HttpClient client = new HttpClient();

        static async Task Main(string[] args)
        {

            var starttime = DateTime.Now;

            var repositories = await ProcessJSON();


            string path = Directory.GetCurrentDirectory();
            string target = @"spreadsheetId.cs";
            Console.WriteLine("The current directory is {0}", path);
            if (!Directory.Exists(target))
            {
                Directory.CreateDirectory(target);
            }

            // Change the current directory.
            Environment.CurrentDirectory = (target);



            using (StreamWriter w = new System.IO.StreamWriter("QUOTES.csv"))
            {

                w.WriteLine("dateTime," + starttime);

                w.WriteLine("BTC," + repositories.BTC.USD);
                w.WriteLine("ETH," + repositories.ETH.USD);
                w.WriteLine("ARMOR," + repositories.ARMOR.USD);
                w.WriteLine("DSLA," + repositories.DSLA.USD);
                w.WriteLine("SFI," + repositories.SFI.USD);
                w.WriteLine("REN," + repositories.REN.USD);
                w.WriteLine("DPI," + repositories.DPI.USD);
                w.WriteLine("ENJ," + repositories.ENJ.USD);

                var oblist = new List<object>() { 12, 3, 4, 5u, 6 };
                Upd("", oblist, repositories);

            }
            var filename = "QUOTES_" + starttime.DayOfWeek.ToString() + "_" + starttime.DayOfYear.ToString() + "_" + starttime.Hour.ToString() + ".csv";
            writeIt(starttime, repositories, target + filename);
            var endtime = DateTime.Now;

        }

        private static void writeIt
            (DateTime starttime, Rootobject2 repositories, string filename)
        {
            using (StreamWriter w = new System.IO.StreamWriter(filename))
            {

                w.WriteLine("dateTime," + starttime);

                w.WriteLine("BTC," + repositories.BTC.USD);
                w.WriteLine("ETH," + repositories.ETH.USD);
                w.WriteLine("ARMOR," + repositories.ARMOR.USD);
                w.WriteLine("DSLA," + repositories.DSLA.USD);
                w.WriteLine("SFI," + repositories.SFI.USD);
                w.WriteLine("REN," + repositories.REN.USD);
                w.WriteLine("DPI," + repositories.DPI.USD);
                w.WriteLine("ENJ," + repositories.ENJ.USD);
                w.WriteLine("XTZ," + repositories.XTZ.USD);
            }
        }

        private static async Task<List<Repository>> ProcessRepositories()
        {
            var uriOriginal = "https://api.github.com/orgs/dotnet/repos";
            var httpHeader = "application/json";  // application/vnd.github.v3+json

            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue(httpHeader));
            client.DefaultRequestHeaders.Add("User-Agent", ".");

            var streamTask = client.GetStreamAsync(uriOriginal);
            var repositories = await JsonSerializer.DeserializeAsync<List<Repository>>(await streamTask);
            return repositories;
        }
        private static async Task<Rootobject2> ProcessJSON()
        {
            // try
            {
                //DEGEN
                var uriCrypto = "https://min-api.cryptocompare.com/data/pricemulti?fsyms=BTC,ETH,ARMOR,DSLA,SFI,REN,DPI,XTZ,ENJ&tsyms=USD"; var httpHeader = "application/json";  // application/vnd.github.v3+json

                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(httpHeader));
                client.DefaultRequestHeaders.Add("User-Agent", ".");

                var streamTask = client.GetStreamAsync(uriCrypto);
                var repositories = await JsonSerializer.DeserializeAsync<Rootobject2>(await streamTask);
                return repositories;

            }
            //catch
            {

            }
        }
        public static UpdateValuesResponse InsertColumnLine(this SheetsService service, string spreadsheetId, string range, params object[] columnValues)
        {
            // convert columnValues to columList
            var columList = columnValues.Select(v => new List<object> { v });

            // Add columList to values and input to valueRange
            var values = new List<IList<object>>();
            values.AddRange(columList.ToList());
            var valueRange = new ValueRange()
            {
                Values = values
            };

            // Create request and execute
            var UpdateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            UpdateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return UpdateRequest.Execute();
        }
        static public void Upd(string celltext, List<object> oblist, Rootobject2 repositories)
        {

            try
            {
                string[] Scopes = { SheetsService.Scope.Spreadsheets, SheetsService.Scope.Drive };
                string ApplicationName = "ssrjap";

                UserCredential credential;
                using (FileStream stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                GoogleClientSecrets.FromStream(stream).Secrets,
                                Scopes,
                                "user",
                                CancellationToken.None).Result; //,
                                                                //new FileDataStore(credPath, true)).Result;
                    //Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                SheetsService service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                String spreadsheetId = "spreadsheetId.cs";
               // String range = "Sheet1!A1:B10";


                //ValueRange valueRange = new ValueRange();
                //valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

                //var oblist = new List<object>() { celltext };
                // IList<IList<object>> list = new List<IList<object>>(); // Works!
                //valueRange.Values =  oblist;

                //var x = repositories.FirstOrDefault();

                InsertColumnLine(service, spreadsheetId, "Sheet1!A1", "DT","BTC", "ETH","ARMOR","DSLA","SFI","REN","DPI","ENJ","XTZ");
                InsertColumnLine(service, spreadsheetId, "Sheet1!B1", DateTime.Now, repositories.BTC.USD, repositories.ETH.USD, repositories.ARMOR.USD, repositories.DSLA.USD, repositories.SFI.USD, repositories.REN.USD, repositories.DPI.USD, repositories.ENJ.USD, repositories.XTZ.USD);

                //SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                //update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                //UpdateValuesResponse result2 = update.Execute();

                // Console.WriteLine("done!");
            }
            catch (Exception ex)
            {

                //                Google.Apis.Requests.RequestError
                //                Request had insufficient authentication scopes. [403]
                //Errors[
                //    Message[Insufficient Permission] Location[- ] Reason[insufficientPermissions] Domain[global]
                //Console.WriteLine(ex.Message.ToString());
            }
        }


    }
    public class Repository
    {
        [JsonPropertyName("BTC")]
        public string BTC { get; set; }

    }

    class Quotes
    {
        public string Symbol { get; set; }
        public Dictionary<string, List<Info>> Quote { get; set; }
    }

    class Info
    {
        public string USD { get; set; }
        public decimal Price { get; set; }
    }

    public class QT { public BTC BTC { get; set; } }


    public class Rootobject2
    {


        public BTC BTC { get; set; }
        public BTC ETH { get; set; }
        public BTC ARMOR { get; set; }
        public BTC DSLA { get; set; }
        public BTC SFI { get; set; }
        public BTC REN { get; set; }
        public BTC DPI { get; set; }
        public BTC ENJ { get; set; }
        public BTC XTZ { get; set; }
    }
    public class BTC
    {
        public float USD { get; set; }
    }


}
