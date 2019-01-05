using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using GSpreadSheet;

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
            else
            {
                Console.WriteLine("{0}", string.Join("\n", executionResult.Messages));
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
            Type cellAddressType = testAssembly.GetType("GSpreadSheet.CellAddress", true);
            Type cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue", true);
            Type executionResultWitDataType = testAssembly.GetType("GSpreadSheet.ExecutionResultWithData`1", true);

            Type listType = typeof(List<>);
            Type[] typeArgs = { cellAddressType };
            Type listGenericType = listType.MakeGenericType(typeArgs);
            object listGenericInstance = Activator.CreateInstance(listGenericType);
            object instanceCellAddress = Activator.CreateInstance(cellAddressType, "", "A8");
            MethodInfo methodAdd = listGenericType.GetMethod("Add");
            methodAdd.Invoke(listGenericInstance, new object[] { instanceCellAddress });

            //  Calling a ReadCellValues function
            //MethodInfo methodReadCellValues = googleSheetsType.GetMethod("ReadCellValues");
            //executionResult = methodReadCellValues.Invoke(gsInstance, new object[] { session, listGenericInstance });
            //bool hasErrors = DisplayErrors(executionResult);
            //if (!hasErrors)
            //{
            //    int i = 1;
            //    foreach (dynamic row in executionResult.Data)
            //    {
            //        var address = row.Address;
            //        var sheetName = row.SheetName;
            //        var val = row.Value;
            //        Console.WriteLine("{0}) - {1}, {2}, {3}", i++, sheetName, address, val);
            //    }
            //}

            //  Creating a CellAddressWithValue instance
            typeArgs[0] = cellAddressValueType;
            listGenericType = listType.MakeGenericType(typeArgs);
            methodAdd = listGenericType.GetMethod("Add");
            listGenericInstance = Activator.CreateInstance(listGenericType);
            object instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType, "", "H5", "Andrew");
            methodAdd.Invoke(listGenericInstance, new object[] { instanceCellAddressValue });
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType, "", "H7", "Constatine");
            methodAdd.Invoke(listGenericInstance, new object[] { instanceCellAddressValue });
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType, "", "H9", "Ricardo");
            methodAdd.Invoke(listGenericInstance, new object[] { instanceCellAddressValue });

            //  Calling a WriteCellValues function
            //MethodInfo methodWriteCellValues = googleSheetsType.GetMethod("WriteCellValues");
            //executionResult = methodWriteCellValues.Invoke(gsInstance, new object[] { session, listGenericInstance });
            //DisplayErrors(executionResult);

            Console.ReadLine();
        }
    }
}
