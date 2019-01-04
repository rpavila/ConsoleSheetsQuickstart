using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ConsoleSheetsQuickstart
{
    class Program
    {        

        static bool DisplayErrors(dynamic executionResult)
        {
            var resultType = (int)executionResult.Result;
            var hasErrors = resultType > 0;
            if (hasErrors)
            {
                Console.WriteLine("{0}) - {1}", resultType > 1 ? "Error" : "Warning", string.Join("\n", executionResult.Messages));
            }
            return hasErrors;
        }

        static void Main(string[] args)
        {
            var currentAssembly = Assembly.GetExecutingAssembly();
            var absolutePath = Path.GetDirectoryName(currentAssembly.Location);
            var testAssembly = Assembly.LoadFile(absolutePath + "/GSpreadSheet.dll");
            var googleSheetsType = testAssembly.GetType("GSpreadSheet.GoogleSheets");

            var constructorParams = new object[] { "credentials.json" };
            var gsInstance = Activator.CreateInstance(googleSheetsType, constructorParams);

            //  Spreadsheets DocID
            var spreadsheetId = "1mfFHzoYsz9Rfypme3bRZGHYqXLVFJNLVz4hNiSS9tfk";
            
            //  Creating a Session
            var methodOpenSession = googleSheetsType.GetMethod("OpenSession");
            var methodReadCellValues = googleSheetsType.GetMethod("ReadCellValues");
            var methodWriteCellValues = googleSheetsType.GetMethod("WriteCellValues");

            var parametersArray = new object[] { spreadsheetId };
            var session = methodOpenSession.Invoke(gsInstance, parametersArray);

            //  Creating a CellAddress instance
            var cellAddressType = testAssembly.GetType("GSpreadSheet.CellAddress");
            var cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue");

//            Type executionResultWitDataType = testAssembly.GetType("GSpreadSheet.ExecutionResultWithData`1", true);

//            FieldInfo propertyAddress = cellAddressType.GetField("Address");


            var listType = typeof(List<>);
            var cellAddressesListType = listType.MakeGenericType(cellAddressType);
            var cellAddressWithValueListType = listType.MakeGenericType(cellAddressValueType);

//            Type[] typeArgs = { typeof(List<object>) };
//            Type makeme = executionResultWitDataType.MakeGenericType(typeArgs);
            dynamic listCellAddress = Activator.CreateInstance(cellAddressesListType);

            dynamic instanceCellAddress = Activator.CreateInstance(cellAddressType, "A2");
//            propertyAddress.SetValue(instanceCellAddress, "A2");
            listCellAddress.Add(instanceCellAddress);

            //  Calling a ReadCellValues function
//            parametersArray = new object[] { session, listCellAddress };
            dynamic executionResult = methodReadCellValues.Invoke(gsInstance, new [] { session, listCellAddress });
            bool hasErrors = DisplayErrors(executionResult);
            if (!hasErrors)
            {
//                IList<object> values = (IList<object>);
                var i = 1;
                foreach (var row in executionResult.Data)
                {
                    var address = row.Address;
                    var sheetName = row.SheetName;
                    var val = row.Value;
                    Console.WriteLine("{0}) - {1}, {2}, {3}", i++, sheetName, address, val);
                }
            }

            //  Creating a CellAddressWithValue instance
            listCellAddress = Activator.CreateInstance(cellAddressWithValueListType);

            dynamic instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType, "H5", "Andrew");
//            propertyAddress = cellAddressValueType.GetField("Address");
//            propertyAddress.SetValue(instanceCellAddressValue, "H5");
//            propertyAddress = cellAddressValueType.GetField("Value");
//            propertyAddress.SetValue(instanceCellAddressValue, "Andrew");
            listCellAddress.Add(instanceCellAddressValue);
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType, "H7", "Constantine");
//            propertyAddress = cellAddressValueType.GetField("Address");
//            propertyAddress.SetValue(instanceCellAddressValue, "H7");
//            propertyAddress = cellAddressValueType.GetField("Value");
//            propertyAddress.SetValue(instanceCellAddressValue, "Constantine");
            listCellAddress.Add(instanceCellAddressValue);
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType, "H9", "Ricardo");
//            propertyAddress = cellAddressValueType.GetField("Address");
//            propertyAddress.SetValue(instanceCellAddressValue, "H9");
//            propertyAddress = cellAddressValueType.GetField("Value");
//            propertyAddress.SetValue(instanceCellAddressValue, "Ricardo");
            listCellAddress.Add(instanceCellAddressValue);

            //  Calling a WriteCellValues function
//            parametersArray = new object[] { session, listCellAddress };
            executionResult = methodWriteCellValues.Invoke(gsInstance, new [] { session, listCellAddress });
            DisplayErrors(executionResult);
            Console.WriteLine("SetValues executed with result={0}", executionResult.Result);
//            Console.ReadLine();
        }
    }
}
