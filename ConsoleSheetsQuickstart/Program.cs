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
            
            var session = methodOpenSession.Invoke(gsInstance, new object[] { spreadsheetId });

            //  Creating a CellAddress instance
            var cellAddressType = testAssembly.GetType("GSpreadSheet.CellAddress", true);
            var cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue", true);

            var listType = typeof(List<>);
            Type[] typeArgs = { cellAddressType };
            var listGenericType = listType.MakeGenericType(typeArgs);
            var listGenericInstance = Activator.CreateInstance(listGenericType);
            var instanceCellAddress = Activator.CreateInstance(cellAddressType, "", "A8");
            var methodAdd = listGenericType.GetMethod("Add");
            methodAdd.Invoke(listGenericInstance, new object[] { instanceCellAddress });

            //  Calling a ReadCellValues function
            dynamic executionResult = methodReadCellValues.Invoke(gsInstance, new object[] { session, listGenericInstance });
            bool hasErrors = DisplayErrors(executionResult);
            if (!hasErrors)
            {
                int i = 1;
                foreach (dynamic row in executionResult.Data)
                {
                    Console.WriteLine("{0}) - {1}, {2}, {3}", i++, row.SheetName, row.Address, row.Value);
                }
            }

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
            //executionResult = methodWriteCellValues.Invoke(gsInstance, new object[] { session, listGenericInstance });
            //DisplayErrors(executionResult);

            Console.ReadLine();
        }
    }
}
