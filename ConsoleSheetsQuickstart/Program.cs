using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Reflection;

namespace ConsoleSheetsQuickstart
{
    class Program
    {        

        static bool DisplayErrors(dynamic executionResult)
        {
            int resultType = (int)executionResult.Result;
            bool hasErrors = resultType > 0;
            if (hasErrors)
            {
                Console.WriteLine("{0}) - {1}", resultType > 1 ? "Error" : "Warning", string.Join("\n", executionResult.Messages));
            }
            return hasErrors;
        }

        static void Main(string[] args)
        {
            Assembly currentAssembly = Assembly.GetExecutingAssembly();
            string absolutePath = Path.GetDirectoryName(currentAssembly.Location);
            Assembly testAssembly = Assembly.LoadFile(absolutePath + "/GSpreadSheet.dll");
            Type googleSheetsType = testAssembly.GetType("GSpreadSheet.GoogleSheets");

            object[] constructorParams = new object[] { "credentials.json" };
            object gsInstance = Activator.CreateInstance(googleSheetsType, constructorParams);

            //  Spreadsheets DocID
            String spreadsheetId = "1mfFHzoYsz9Rfypme3bRZGHYqXLVFJNLVz4hNiSS9tfk";
            
            //  Creating a Session
            MethodInfo methodOpenSession = googleSheetsType.GetMethod("OpenSession");
            object[] parametersArray = new object[] { spreadsheetId };
            object session = methodOpenSession.Invoke(gsInstance, parametersArray);

            //  Creating a CellAddress instance
            Type cellAddressType = testAssembly.GetType("GSpreadSheet.CellAddress");
            Type cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue");
            Type executionResultWitDataType = testAssembly.GetType("GSpreadSheet.ExecutionResultWithData`1", true);
            Type[] typeArgs = { typeof(List<object>) };
            Type makeme = executionResultWitDataType.MakeGenericType(typeArgs);
            List<object> listCellAddress = new List<object>();

            object instanceCellAddress = Activator.CreateInstance(cellAddressType);
            FieldInfo propertyAddress = cellAddressType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddress, "A2");
            listCellAddress.Add(instanceCellAddress);

            dynamic executionResult = null;
            //  Calling a ReadCellValues function
            //MethodInfo methodReadCellValues = googleSheetsType.GetMethod("ReadCellValues");
            //parametersArray = new object[] { session, listCellAddress };
            //executionResult = methodReadCellValues.Invoke(gsInstance, parametersArray);
            //bool hasErrors = DisplayErrors(executionResult);
            //if (!hasErrors)
            //{
            //    IList<object> values = (IList<object>)executionResult.Data;
            //    int i = 1;
            //    foreach (dynamic row in values)
            //    {
            //        var address = row.Address;
            //        var sheetName = row.SheetName;
            //        var val = row.Value;
            //        Console.WriteLine("{0}) - {1}, {2}, {3}", i++, sheetName, address, val);
            //    }
            //}

            //  Creating a CellAddressWithValue instance
            listCellAddress = new List<object>();
            cellAddressValueType = testAssembly.GetType("GSpreadSheet.CellAddressWithValue");
            object instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType);
            propertyAddress = cellAddressValueType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddressValue, "H5");
            propertyAddress = cellAddressValueType.GetField("Value");
            propertyAddress.SetValue(instanceCellAddressValue, "Andrew");
            listCellAddress.Add(instanceCellAddressValue);
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType);
            propertyAddress = cellAddressValueType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddressValue, "H7");
            propertyAddress = cellAddressValueType.GetField("Value");
            propertyAddress.SetValue(instanceCellAddressValue, "Constantine");
            listCellAddress.Add(instanceCellAddressValue);
            instanceCellAddressValue = Activator.CreateInstance(cellAddressValueType);
            propertyAddress = cellAddressValueType.GetField("Address");
            propertyAddress.SetValue(instanceCellAddressValue, "H9");
            propertyAddress = cellAddressValueType.GetField("Value");
            propertyAddress.SetValue(instanceCellAddressValue, "Ricardo");
            listCellAddress.Add(instanceCellAddressValue);

            //  Calling a WriteCellValues function
            //MethodInfo methodWriteCellValues = googleSheetsType.GetMethod("WriteCellValues");
            //parametersArray = new object[] { session, listCellAddress };
            //executionResult = methodWriteCellValues.Invoke(gsInstance, parametersArray);
            //DisplayErrors(executionResult);

            Console.ReadLine();
        }
    }
}
