using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Reflection;

namespace ConsoleSheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        //static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        //static string ApplicationName = "Google Sheets API .NET Quickstart";
        
        static void Main(string[] args)
        {
            Assembly testAssembly = Assembly.LoadFile(@"c:\GSpreadSheet.dll");

            // get type of class Calculator from just loaded assembly
            Type googleSheetsType = testAssembly.GetType("GSpreadSheet.GoogleSheets");

            // create instance of class Calculator
            object gsInstance = Activator.CreateInstance(googleSheetsType);

            //String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String spreadsheetId = "1mfFHzoYsz9Rfypme3bRZGHYqXLVFJNLVz4hNiSS9tfk";

            //GoogleSheets gs = new GoogleSheets();
            //object doc = gs.OpenSession(spreadsheetId);
            MethodInfo methodOpenSession = googleSheetsType.GetMethod("OpenSession");
            object[] parametersArray = new object[] { spreadsheetId };
            object session = methodOpenSession.Invoke(gsInstance, parametersArray);


            //List<CellAddress> Values = new List<CellAddress> {
            //    new CellAddress{ Address = "A2:E"}
            //};
            //IList<CellAddressWithValue> values = gs.ReadCellValues(doc, Values);

            Type cellAddressType = testAssembly.GetType("GSpreadSheet.CellAddress");
            Type cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue");
            //object[] listCellAddress = (object[])Activator.CreateInstance(cellAddressType);
            List<object> listCellAddress = new List<object>();

            object instanceCellAddress = Activator.CreateInstance(cellAddressType);
            FieldInfo propertyAddress = cellAddressType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddress, "A2:I");
            //propertyAddress = cellAddressType.GetProperty("SheetName");
            //propertyAddress.SetValue(instanceCellAddress, "Class Data");

            listCellAddress.Add(instanceCellAddress);

            MethodInfo methodReadCellValues = googleSheetsType.GetMethod("ReadCellValues");
            parametersArray = new object[] { session, listCellAddress };
            IList<object> values = (IList<object>)methodReadCellValues.Invoke(gsInstance, parametersArray);

            int i = 1;
            foreach (var row in values)
            {
                propertyAddress = cellAddressValueType.GetField("Address");
                var address = propertyAddress.GetValue(row);
                propertyAddress = cellAddressValueType.GetField("SheetName");
                var sheetName = propertyAddress.GetValue(row);
                propertyAddress = cellAddressValueType.GetField("Value");
                var val = propertyAddress.GetValue(row);
                Console.WriteLine("{0}) - {1}, {2}, {3}", i++, sheetName, address, val);
            }

            //List<CellAddressWithValue> ValuesWrite = new List<CellAddressWithValue> {
            //    new CellAddressWithValue{ Address = "A2", Value = "Hello world!!!"},
            //    new CellAddressWithValue{ Address = "C5", Value = "Hello Ricardo!!!"},
            //    new CellAddressWithValue{ Address = "E8", Value = "Hello Andrew!!!"},
            //    new CellAddressWithValue{ Address = "J8", Value = "Hello Andrew!!!"}
            //};
            //gs.WriteCellValues(doc, ValuesWrite);

            Console.ReadLine();

            //UserCredential credential;

            //using (var stream =
            //    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            //{
            //    // The file token.json stores the user's access and refresh tokens, and is created
            //    // automatically when the authorization flow completes for the first time.
            //    string credPath = "token.json";
            //    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            //        GoogleClientSecrets.Load(stream).Secrets,
            //        Scopes,
            //        "user",
            //        CancellationToken.None,
            //        new FileDataStore(credPath, true)).Result;
            //    Console.WriteLine("Credential file saved to: " + credPath);
            //}

            //// Create Google Sheets API service.
            //var service = new SheetsService(new BaseClientService.Initializer()
            //{
            //    HttpClientInitializer = credential,
            //    ApplicationName = ApplicationName,
            //});

            //// Define request parameters.
            //String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            //String range = "Class Data!A2:E";
            //SpreadsheetsResource.ValuesResource.GetRequest request =
            //        service.Spreadsheets.Values.Get(spreadsheetId, range);

            //// Prints the names and majors of students in a sample spreadsheet:
            //// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            //ValueRange response = request.Execute();
            //IList<IList<Object>> values = response.Values;
            //if (values != null && values.Count > 0)
            //{
            //    Console.WriteLine("Name, Major");
            //    foreach (var row in values)
            //    {
            //        // Print columns A and E, which correspond to indices 0 and 4.
            //        Console.WriteLine("{0}, {1}", row[0], row[4]);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No data found.");
            //}
            //Console.Read();
        }
    }
}
